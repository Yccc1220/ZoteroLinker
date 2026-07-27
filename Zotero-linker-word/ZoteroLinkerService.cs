using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text.RegularExpressions;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

namespace Zotero_linker
{
    internal sealed class ZoteroLinkerService
    {
        private const float DefaultCitationFontSize = 10f;

        internal LinkResult LinkCitations(Word.Document document, Color citationColor, float fontSize)
        {
            if (document == null)
            {
                throw new InvalidOperationException("No active Word document.");
            }

            bool previousScreenUpdating = document.Application.ScreenUpdating;
            document.Application.ScreenUpdating = false;
            try
            {
                EnsureZoteroHyperlinkStyles(document, citationColor);
                RemoveCitationLinks(document);

                ZoteroFields fields = LoadZoteroFields(document);
                LinkResult result = new LinkResult();
                if (fields.Bibliography == null)
                {
                    result.SkippedMissingBibliography = fields.Citations.Count;
                    return result;
                }

                HashSet<string> backlinkedItems = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                List<PendingBibliographyBacklink> pendingBibliographyBacklinks = new List<PendingBibliographyBacklink>();
                foreach (ZoteroFieldInfo citation in fields.Citations)
                {
                    List<CitationSegment> segments = ExtractCitationSegments(citation);
                    if (segments.Count == 0)
                    {
                        result.SkippedMultiItem += 1;
                        continue;
                    }

                    result.SkippedCompressedItems += segments.Count(segment => !segment.Visible);
                    List<PendingCitationLink> pendingCitationLinks = new List<PendingCitationLink>();

                    foreach (CitationSegment segment in segments)
                    {
                        Word.Range bibliographyRange = FindBibliographyEntryRange(fields.Bibliography, segment.Item);
                        if (bibliographyRange == null)
                        {
                            result.FailedBibliographyMatch += 1;
                            continue;
                        }

                        string bibliographyBookmarkName = BuildBookmarkName(segment.Item);
                        string citationBookmarkName = BuildCitationBookmarkName(segment.Item);
                        CitationSegment backlinkSourceSegment = GetCitationBacklinkSourceSegment(segment);
                        Word.Range citationBacklinkRange = RangeFromOffsets(
                            document,
                            citation.Field.Result,
                            backlinkSourceSegment.StartOffset,
                            backlinkSourceSegment.EndOffset);

                        if (citationBacklinkRange == null)
                        {
                            result.FailedCitationRange += 1;
                            continue;
                        }

                        Word.Range citationRange = null;
                        if (segment.Visible)
                        {
                            citationRange = RangeFromOffsets(document, citation.Field.Result, segment.StartOffset, segment.EndOffset);
                            if (citationRange == null)
                            {
                                result.FailedCitationRange += 1;
                                continue;
                            }
                        }

                        AddOrReplaceBookmark(document, bibliographyBookmarkName, bibliographyRange);
                        if (!backlinkedItems.Contains(bibliographyBookmarkName))
                        {
                            AddOrReplaceBookmark(document, citationBookmarkName, citationBacklinkRange);
                            pendingBibliographyBacklinks.Add(new PendingBibliographyBacklink(
                                bibliographyBookmarkName,
                                citationBookmarkName));
                            backlinkedItems.Add(bibliographyBookmarkName);
                        }

                        if (citationRange != null)
                        {
                            string temporaryBookmarkName = BuildTemporaryCitationBookmarkName(pendingCitationLinks.Count);
                            AddOrReplaceBookmark(document, temporaryBookmarkName, citationRange);
                            ApplyCitationFormatting(citationRange, citationColor, fontSize);
                            pendingCitationLinks.Add(new PendingCitationLink(
                                temporaryBookmarkName,
                                bibliographyBookmarkName,
                                SegmentTooltip(segment.Item, bibliographyBookmarkName)));
                        }
                    }

                    foreach (PendingCitationLink pendingLink in pendingCitationLinks)
                    {
                        if (!document.Bookmarks.Exists(pendingLink.TemporaryBookmarkName))
                        {
                            result.FailedCitationRange += 1;
                            continue;
                        }

                        Word.Range temporaryRange = document.Bookmarks[pendingLink.TemporaryBookmarkName].Range;
                        Word.Hyperlink citationHyperlink = AddInternalHyperlink(
                            document,
                            temporaryRange,
                            pendingLink.BibliographyBookmarkName,
                            pendingLink.ScreenTip);
                        ApplyCitationFormatting(citationHyperlink.Range, citationColor, fontSize);
                        ApplyCitationFormatting(temporaryRange, citationColor, fontSize);
                        document.Bookmarks[pendingLink.TemporaryBookmarkName].Delete();
                        result.Linked += 1;
                    }
                }

                result.LinkedBacklinks += AddPendingBibliographyBacklinks(document, pendingBibliographyBacklinks);
                RestoreCitationFormatting(document, citationColor, fontSize);
                return result;
            }
            finally
            {
                document.Application.ScreenUpdating = previousScreenUpdating;
            }
        }

        internal RemoveResult RemoveCitationLinks(Word.Document document)
        {
            if (document == null)
            {
                throw new InvalidOperationException("No active Word document.");
            }

            RemoveResult result = new RemoveResult();
            HashSet<string> legacyBookmarkNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            List<Word.Range> citationRanges = GetZoteroCitationResultRanges(document);

            for (int index = document.Hyperlinks.Count; index >= 1; index -= 1)
            {
                Word.Hyperlink hyperlink = document.Hyperlinks[index];
                string subAddress = NormalizeCiteName(hyperlink.SubAddress);
                bool isCurrentLink = IsCiteName(subAddress);
                bool isLegacyMacroLink = !isCurrentLink &&
                    IsLegacyMacroBookmarkName(subAddress) &&
                    RangeOverlapsAny(hyperlink.Range, citationRanges);

                if (!isCurrentLink && !isLegacyMacroLink)
                {
                    continue;
                }

                if (isCurrentLink && IsCitationBackBookmarkName(subAddress))
                {
                    ApplyBibliographyBacklinkFormatting(hyperlink.Range);
                }

                if (isLegacyMacroLink)
                {
                    legacyBookmarkNames.Add(subAddress);
                }

                hyperlink.Delete();
                result.LinksRemoved += 1;
            }

            for (int index = document.Bookmarks.Count; index >= 1; index -= 1)
            {
                Word.Bookmark bookmark = document.Bookmarks[index];
                if (IsCiteName(bookmark.Name) || legacyBookmarkNames.Contains(bookmark.Name))
                {
                    bookmark.Delete();
                    result.BookmarksRemoved += 1;
                }
            }

            foreach (Word.Field field in document.Fields)
            {
                if (IsZoteroCitationCode(SafeFieldCode(field)))
                {
                    if (ResetZoteroFieldFormatting(field.Result))
                    {
                        result.Recolored += 1;
                    }
                }
            }

            return result;
        }

        internal int RestoreCitationFormatting(Word.Document document, Color citationColor, float fontSize)
        {
            if (document == null)
            {
                throw new InvalidOperationException("No active Word document.");
            }

            EnsureZoteroHyperlinkStyles(document, citationColor);

            int changed = 0;
            foreach (Word.Field field in document.Fields)
            {
                if (IsZoteroCitationCode(SafeFieldCode(field)))
                {
                    if (ApplyCitationFormatting(field.Result, citationColor, fontSize))
                    {
                        changed += 1;
                    }
                }
            }

            foreach (Word.Hyperlink hyperlink in document.Hyperlinks)
            {
                string subAddress = NormalizeCiteName(hyperlink.SubAddress);
                if (!IsCiteName(subAddress))
                {
                    continue;
                }

                if (IsCitationBackBookmarkName(subAddress))
                {
                    ApplyBibliographyBacklinkFormatting(hyperlink.Range);
                }
                else
                {
                    ApplyCitationFormatting(hyperlink.Range, citationColor, fontSize);
                }
            }

            return changed;
        }

        internal void EnsureCitationLinkStyles(Word.Document document, Color citationColor)
        {
            if (document == null)
            {
                throw new InvalidOperationException("No active Word document.");
            }

            EnsureZoteroHyperlinkStyles(document, citationColor);
        }

        internal int RestoreLinkedCitationFormatting(Word.Document document, Color citationColor, float fontSize)
        {
            if (document == null)
            {
                throw new InvalidOperationException("No active Word document.");
            }

            EnsureZoteroHyperlinkStyles(document, citationColor);

            int changed = 0;
            foreach (Word.Hyperlink hyperlink in document.Hyperlinks)
            {
                string subAddress = NormalizeCiteName(hyperlink.SubAddress);
                if (!IsCiteName(subAddress))
                {
                    continue;
                }

                if (IsCitationBackBookmarkName(subAddress))
                {
                    if (ApplyBibliographyBacklinkFormatting(hyperlink.Range))
                    {
                        changed += 1;
                    }
                }
                else
                {
                    if (ApplyCitationFormatting(hyperlink.Range, citationColor, fontSize))
                    {
                        changed += 1;
                    }
                }
            }

            return changed;
        }

        internal void RestoreFollowedHyperlinkFormatting(Word.Hyperlink hyperlink, Color citationColor, float fontSize)
        {
            if (hyperlink == null)
            {
                return;
            }

            string subAddress = NormalizeCiteName(hyperlink.SubAddress);
            if (!IsCiteName(subAddress))
            {
                return;
            }

            if (IsCitationBackBookmarkName(subAddress))
            {
                ApplyBibliographyBacklinkFormatting(hyperlink.Range);
            }
            else
            {
                ApplyCitationFormatting(hyperlink.Range, citationColor, fontSize);
            }
        }

        private static ZoteroFields LoadZoteroFields(Word.Document document)
        {
            ZoteroFields fields = new ZoteroFields();
            foreach (Word.Field field in document.Fields)
            {
                string code = SafeFieldCode(field);
                if (IsZoteroCitationCode(code))
                {
                    fields.Citations.Add(new ZoteroFieldInfo
                    {
                        Field = field,
                        Code = code,
                        Text = field.Result.Text ?? string.Empty,
                        CitationItems = ParseCitationItems(code)
                    });
                }
                else if (IsZoteroBibliographyCode(code))
                {
                    fields.Bibliography = field;
                }
            }

            return fields;
        }

        private static List<Word.Range> GetZoteroCitationResultRanges(Word.Document document)
        {
            List<Word.Range> ranges = new List<Word.Range>();
            if (document == null)
            {
                return ranges;
            }

            foreach (Word.Field field in document.Fields)
            {
                if (IsZoteroCitationCode(SafeFieldCode(field)))
                {
                    ranges.Add(field.Result);
                }
            }

            return ranges;
        }

        private static Word.Range FindBibliographyEntryRange(Word.Field bibliographyField, CitationItem item)
        {
            if (bibliographyField == null || item == null)
            {
                return null;
            }

            string bibliographyText = bibliographyField.Result.Text ?? string.Empty;
            BibliographyEntryMatch bestMatch = null;
            foreach (BibliographyEntryMatch entry in EnumerateBibliographyEntries(bibliographyText))
            {
                entry.Score = ScoreBibliographyEntry(entry.Text, item);
                if (entry.Score <= 0)
                {
                    continue;
                }

                if (bestMatch == null || entry.Score > bestMatch.Score)
                {
                    bestMatch = entry;
                }
            }

            if (bestMatch != null)
            {
                return bibliographyField.Result.Document.Range(
                    bibliographyField.Result.Start + bestMatch.StartOffset,
                    bibliographyField.Result.Start + bestMatch.EndOffset);
            }

            return null;
        }

        private static IEnumerable<BibliographyEntryMatch> EnumerateBibliographyEntries(string bibliographyText)
        {
            string text = bibliographyText ?? string.Empty;
            int startOffset = 0;
            while (startOffset < text.Length)
            {
                int endOffset = text.IndexOf('\r', startOffset);
                if (endOffset < 0)
                {
                    endOffset = text.Length;
                }

                int trimmedStart = startOffset;
                int trimmedEnd = endOffset;
                while (trimmedStart < trimmedEnd && char.IsWhiteSpace(text[trimmedStart]))
                {
                    trimmedStart += 1;
                }

                while (trimmedEnd > trimmedStart && char.IsWhiteSpace(text[trimmedEnd - 1]))
                {
                    trimmedEnd -= 1;
                }

                if (trimmedStart < trimmedEnd)
                {
                    yield return new BibliographyEntryMatch(
                        trimmedStart,
                        trimmedEnd,
                        text.Substring(trimmedStart, trimmedEnd - trimmedStart));
                }

                startOffset = endOffset + 1;
            }
        }

        private static int ScoreBibliographyEntry(string entryText, CitationItem item)
        {
            string entry = NormalizeBibliographyMatchText(entryText);
            if (string.IsNullOrWhiteSpace(entry) || item == null)
            {
                return 0;
            }

            string title = NormalizeBibliographyMatchText(ShortenSearchText(item.Title ?? string.Empty, 120));
            string author = NormalizeBibliographyMatchText(item.Author);
            string year = NormalizeBibliographyMatchText(item.Year);

            bool titleMatch = !string.IsNullOrWhiteSpace(title) && entry.Contains(title);
            bool authorMatch = !string.IsNullOrWhiteSpace(author) && ContainsBibliographyToken(entry, author);
            bool yearMatch = !string.IsNullOrWhiteSpace(year) && ContainsBibliographyToken(entry, year);

            int score = 0;
            if (titleMatch)
            {
                score += title.Length >= 30 ? 60 : title.Length >= 15 ? 40 : 25;
            }

            if (authorMatch)
            {
                score += author.Length <= 2 ? 18 : 25;
            }

            if (yearMatch)
            {
                score += 15;
            }

            if (titleMatch && authorMatch)
            {
                score += 20;
            }

            if (titleMatch && yearMatch)
            {
                score += 10;
            }

            if (authorMatch && yearMatch)
            {
                score += 10;
            }

            if (titleMatch && !authorMatch && !yearMatch && title.Length < 24)
            {
                score -= 20;
            }

            return Math.Max(0, score);
        }

        private static Word.Range FindBibliographyBacklinkAnchorRange(Word.Range bibliographyRange, CitationItem item)
        {
            string text = bibliographyRange.Text ?? string.Empty;
            string anchorText = ExtractBibliographyAnchorText(text);
            if (!string.IsNullOrWhiteSpace(anchorText))
            {
                int start = bibliographyRange.Start + text.IndexOf(anchorText, StringComparison.Ordinal);
                int end = start + anchorText.Length;
                return bibliographyRange.Document.Range(start, end);
            }

            foreach (string candidate in BuildBibliographyBacklinkAnchorCandidates(text, item))
            {
                Word.Range found = FindFirstRange(bibliographyRange, candidate, false);
                if (found != null)
                {
                    return found;
                }
            }

            return null;
        }

        private static int AddPendingBibliographyBacklinks(
            Word.Document document,
            IEnumerable<PendingBibliographyBacklink> pendingBacklinks)
        {
            List<PendingBibliographyBacklinkRange> ranges = new List<PendingBibliographyBacklinkRange>();
            foreach (PendingBibliographyBacklink pendingBacklink in pendingBacklinks)
            {
                if (!document.Bookmarks.Exists(pendingBacklink.BibliographyBookmarkName) ||
                    !document.Bookmarks.Exists(pendingBacklink.CitationBookmarkName))
                {
                    continue;
                }

                Word.Range bibliographyRange = document.Bookmarks[pendingBacklink.BibliographyBookmarkName].Range;
                Word.Range backlinkRange = FindBibliographyBacklinkAnchorRange(bibliographyRange, null);
                if (backlinkRange == null)
                {
                    continue;
                }

                ranges.Add(new PendingBibliographyBacklinkRange(
                    backlinkRange.Start,
                    backlinkRange.End,
                    pendingBacklink.CitationBookmarkName));
            }

            int linked = 0;
            foreach (PendingBibliographyBacklinkRange pendingRange in ranges.OrderByDescending(range => range.Start))
            {
                Word.Range backlinkRange = document.Range(pendingRange.Start, pendingRange.End);
                Word.Hyperlink backlink = AddInternalHyperlink(
                    document,
                    backlinkRange,
                    pendingRange.CitationBookmarkName,
                    "Back to citation",
                    true);
                ApplyBibliographyBacklinkFormatting(backlink.Range);
                ApplyBibliographyBacklinkFormatting(backlinkRange);
                linked += 1;
            }

            return linked;
        }

        private static string ExtractBibliographyAnchorText(string text)
        {
            Match numericPrefix = Regex.Match(
                text ?? string.Empty,
                @"^\s*((?:[\[\uff3b][0-9,\uff0c;\uff1b\s\-\u2010\u2011\u2012\u2013\u2014]+[\]\uff3d])|(?:[0-9]+(?:[.)]|\u3001|\uff0e|\uff1a|\uff09)?))(?:\s*(?:->|=>|\u2192|\u21d2))?");

            return numericPrefix.Success && numericPrefix.Groups[1].Length > 0
                ? numericPrefix.Groups[1].Value
                : null;
        }

        private static List<string> BuildBibliographyBacklinkAnchorCandidates(string text, CitationItem item)
        {
            List<string> candidates = new List<string>();
            if (item != null && !string.IsNullOrWhiteSpace(item.Author))
            {
                candidates.Add(item.Author.Trim());
            }

            if (item != null && !string.IsNullOrWhiteSpace(item.Title))
            {
                candidates.Add(ShortenSearchText(item.Title.Trim(), 80));
            }

            string leadingText = ShortenSearchText((text ?? string.Empty).Trim(), 40);
            if (!string.IsNullOrWhiteSpace(leadingText))
            {
                candidates.Add(leadingText);
            }

            return candidates.Where(value => !string.IsNullOrWhiteSpace(value))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static List<CitationSegment> ExtractCitationSegments(ZoteroFieldInfo citation)
        {
            if (citation.CitationItems.Count == 0)
            {
                return new List<CitationSegment>();
            }

            if (citation.CitationItems.Count == 1)
            {
                CitationBounds bounds = TrimCitationBounds(citation.Text, 0, citation.Text.Length);
                return new List<CitationSegment>
                {
                    new CitationSegment(citation.CitationItems[0], bounds.StartOffset, bounds.EndOffset, true)
                };
            }

            if (IsLikelyNumericCitation(citation.Text))
            {
                List<CitationSegment> numericSegments = ExtractNumericCitationSegments(citation);
                if (numericSegments.Count > 0 && CoversAllCitationItems(numericSegments, citation.CitationItems.Count))
                {
                    return numericSegments;
                }
            }

            List<CitationSegment> delimitedSegments = ExtractDelimitedCitationSegments(citation);
            if (delimitedSegments.Count > 0 && CoversAllCitationItems(delimitedSegments, citation.CitationItems.Count))
            {
                return delimitedSegments;
            }

            return ExtractAuthorYearCitationSegments(citation);
        }

#if DEBUG
        internal static List<CitationSegment> TestExtractCitationSegments(string citationText, int citationItemCount)
        {
            ZoteroFieldInfo citation = new ZoteroFieldInfo
            {
                Text = citationText ?? string.Empty,
                CitationItems = Enumerable.Range(1, Math.Max(0, citationItemCount))
                    .Select(index => new CitationItem
                    {
                        Id = index.ToString(),
                        Title = "Title " + index,
                        Author = "Author" + index,
                        Year = "20" + index.ToString("00")
                    })
                    .ToList()
            };

            return ExtractCitationSegments(citation);
        }

        internal static string TestBuildBookmarkName(string id, string title, string year, string author)
        {
            return BuildBookmarkName(new CitationItem
            {
                Id = id,
                Title = title,
                Year = year,
                Author = author
            });
        }

        internal static string TestExtractBibliographyAnchorText(string bibliographyText)
        {
            return ExtractBibliographyAnchorText(bibliographyText ?? string.Empty) ?? string.Empty;
        }

        internal static bool TestIsLegacyMacroBookmarkName(string bookmarkName)
        {
            return IsLegacyMacroBookmarkName(bookmarkName);
        }
#endif

        private static List<CitationSegment> ExtractNumericCitationSegments(ZoteroFieldInfo citation)
        {
            List<NumericToken> tokens = TokenizeCitationNumbers(citation.Text);
            List<CitationSegment> segments = new List<CitationSegment>();
            if (tokens.Count == 0)
            {
                return segments;
            }

            int itemIndex = -1;
            RangeStart rangeStart = null;
            CitationSegment rangeStartSegment = null;

            foreach (NumericToken token in tokens)
            {
                if (token.SeparatorAfter == NumericSeparator.Dash)
                {
                    itemIndex += 1;
                    CitationItem item = GetItemAt(citation.CitationItems, itemIndex);
                    if (item != null)
                    {
                        rangeStartSegment = new CitationSegment(item, token.StartOffset, token.EndOffset, true);
                        segments.Add(rangeStartSegment);
                        rangeStart = new RangeStart(itemIndex, token);
                    }

                    continue;
                }

                if (rangeStart != null)
                {
                    int endItemIndex = itemIndex + Math.Max(1, token.Value - rangeStart.Token.Value);
                    int backlinkStartOffset = rangeStart.Token.StartOffset;
                    int backlinkEndOffset = rangeStart.Token.EndOffset;
                    if (rangeStartSegment != null)
                    {
                        rangeStartSegment.BacklinkStartOffset = backlinkStartOffset;
                        rangeStartSegment.BacklinkEndOffset = backlinkEndOffset;
                    }

                    for (int hiddenIndex = rangeStart.ItemIndex + 1; hiddenIndex < endItemIndex; hiddenIndex += 1)
                    {
                        CitationItem hiddenItem = GetItemAt(citation.CitationItems, hiddenIndex);
                        if (hiddenItem != null)
                        {
                            segments.Add(new CitationSegment(hiddenItem, backlinkStartOffset, backlinkEndOffset, false)
                            {
                                BacklinkStartOffset = backlinkStartOffset,
                                BacklinkEndOffset = backlinkEndOffset
                            });
                        }
                    }

                    itemIndex = endItemIndex;
                    CitationItem endItem = GetItemAt(citation.CitationItems, itemIndex);
                    if (endItem != null)
                    {
                        segments.Add(new CitationSegment(endItem, token.StartOffset, token.EndOffset, true)
                        {
                            BacklinkStartOffset = token.StartOffset,
                            BacklinkEndOffset = token.EndOffset
                        });
                    }

                    rangeStart = null;
                    rangeStartSegment = null;
                    continue;
                }

                itemIndex += 1;
                CitationItem normalItem = GetItemAt(citation.CitationItems, itemIndex);
                if (normalItem != null)
                {
                    segments.Add(new CitationSegment(normalItem, token.StartOffset, token.EndOffset, true));
                }
            }

            return segments;
        }

        private static bool IsLikelyNumericCitation(string text)
        {
            string value = NormalizeCitationDisplayText(text);
            if (string.IsNullOrWhiteSpace(value))
            {
                return false;
            }

            if (Regex.IsMatch(value, @"^[\s\[\]\uff3b\uff3d\(\)\{\}\uff08\uff09\u3014\u3015,\uff0c;\uff1b.:\uff1a\uff0e\-\u2010\u2011\u2012\u2013\u2014\d]+$"))
            {
                return true;
            }

            return Regex.IsMatch(value, @"^\s*(?:[\[\uff3b]\s*\d+|[\(\uff08]\s*\d+|\d+\s*(?:[,\uff0c\.;\uff1b:\uff1a\uff0e\-\u2010\u2011\u2012\u2013\u2014]|\z))");
        }

        private static bool CoversAllCitationItems(List<CitationSegment> segments, int citationItemCount)
        {
            if (citationItemCount <= 0 || segments == null || segments.Count == 0)
            {
                return false;
            }

            return segments.Select(segment => segment.Item)
                .Where(item => item != null)
                .Distinct(new CitationItemReferenceComparer())
                .Count() == citationItemCount;
        }

        private static List<NumericToken> TokenizeCitationNumbers(string text)
        {
            List<NumericToken> tokens = new List<NumericToken>();
            foreach (Match match in Regex.Matches(text ?? string.Empty, @"\d+"))
            {
                tokens.Add(new NumericToken
                {
                    Value = int.Parse(match.Value),
                    StartOffset = match.Index,
                    EndOffset = match.Index + match.Length
                });
            }

            foreach (NumericToken token in tokens)
            {
                token.SeparatorAfter = DetectNumericSeparatorAfter(text ?? string.Empty, token.EndOffset);
            }

            return tokens;
        }

        private static NumericSeparator DetectNumericSeparatorAfter(string text, int offset)
        {
            for (int index = offset; index < text.Length; index += 1)
            {
                char character = text[index];
                if (char.IsWhiteSpace(character))
                {
                    continue;
                }

                if (IsClosingCitationWrapper(character) || IsOpeningCitationWrapper(character))
                {
                    continue;
                }

                if (IsRangeDash(character))
                {
                    return NumericSeparator.Dash;
                }

                if (IsCitationSeparator(character))
                {
                    return NumericSeparator.Comma;
                }

                return NumericSeparator.None;
            }

            return NumericSeparator.End;
        }

        private static List<CitationSegment> ExtractDelimitedCitationSegments(ZoteroFieldInfo citation)
        {
            List<CitationBounds> chunks = new List<CitationBounds>();
            int startOffset = 0;
            for (int index = 0; index <= citation.Text.Length; index += 1)
            {
                char character = index < citation.Text.Length ? citation.Text[index] : '\0';
                if (index == citation.Text.Length || character == ';' || character == '\uff1b')
                {
                    CitationBounds bounds = TrimCitationBounds(citation.Text, startOffset, index);
                    if (bounds.StartOffset < bounds.EndOffset)
                    {
                        chunks.Add(bounds);
                    }

                    startOffset = index + 1;
                }
            }

            if (chunks.Count != citation.CitationItems.Count)
            {
                return new List<CitationSegment>();
            }

            List<CitationSegment> segments = new List<CitationSegment>();
            for (int index = 0; index < chunks.Count; index += 1)
            {
                segments.Add(new CitationSegment(
                    citation.CitationItems[index],
                    chunks[index].StartOffset,
                    chunks[index].EndOffset,
                    true));
            }

            return segments;
        }

        private static List<CitationSegment> ExtractAuthorYearCitationSegments(ZoteroFieldInfo citation)
        {
            foreach (Tuple<bool, bool> mode in new[]
            {
                Tuple.Create(false, true),
                Tuple.Create(false, false),
                Tuple.Create(true, true)
            })
            {
                List<CitationSegment> segments = ExtractAuthorYearCitationSegments(
                    citation,
                    mode.Item1,
                    mode.Item2);
                if (segments.Count > 0 && CoversAllCitationItems(segments, citation.CitationItems.Count))
                {
                    return segments;
                }
            }

            return new List<CitationSegment>();
        }

        private static List<CitationSegment> ExtractAuthorYearCitationSegments(
            ZoteroFieldInfo citation,
            bool onlyYear,
            bool multiRefCommaSep)
        {
            List<CitationSegment> segments = new List<CitationSegment>();
            string text = citation.Text ?? string.Empty;
            bool inCitation = false;
            bool beginYear = false;
            int commaCount = 0;
            int startOffset = 0;

            for (int index = 0; index < text.Length; index += 1)
            {
                char character = text[index];
                bool createCitation = false;
                bool restartAfterDelimiter = false;

                if (IsOpeningCitationWrapper(character) && !onlyYear)
                {
                    inCitation = true;
                    startOffset = index + 1;
                }
                else if (char.IsDigit(character))
                {
                    beginYear = true;
                    if (onlyYear && !inCitation)
                    {
                        inCitation = true;
                        startOffset = index;
                    }
                }
                else if (multiRefCommaSep && IsCommaLikeChar(character))
                {
                    commaCount += 1;
                    if (commaCount > 1 && beginYear)
                    {
                        createCitation = true;
                        restartAfterDelimiter = !onlyYear;
                    }
                }
                else if (IsSemicolonLikeChar(character) || IsClosingCitationWrapper(character))
                {
                    beginYear = false;
                    if (multiRefCommaSep)
                    {
                        commaCount = 0;
                    }

                    createCitation = true;
                    restartAfterDelimiter = (IsSemicolonLikeChar(character) || IsCommaLikeChar(character)) && !onlyYear;
                }

                if (!createCitation || !inCitation)
                {
                    continue;
                }

                CitationBounds bounds = TrimCitationBounds(text, startOffset, index);
                AddCitationSegmentIfValid(citation, segments, bounds.StartOffset, bounds.EndOffset);
                inCitation = false;

                if (restartAfterDelimiter)
                {
                    int nextStart = index + 1;
                    while (nextStart < text.Length && char.IsWhiteSpace(text[nextStart]))
                    {
                        nextStart += 1;
                    }

                    if (nextStart < text.Length)
                    {
                        startOffset = nextStart;
                        inCitation = true;
                    }
                    else
                    {
                        startOffset = index;
                    }
                }
            }

            if (inCitation)
            {
                CitationBounds bounds = TrimCitationBounds(text, startOffset, text.Length);
                AddCitationSegmentIfValid(citation, segments, bounds.StartOffset, bounds.EndOffset);
            }

            return segments;
        }

        private static void AddCitationSegmentIfValid(
            ZoteroFieldInfo citation,
            List<CitationSegment> segments,
            int startOffset,
            int endOffset)
        {
            if (startOffset >= endOffset || segments.Count >= citation.CitationItems.Count)
            {
                return;
            }

            segments.Add(new CitationSegment(
                citation.CitationItems[segments.Count],
                startOffset,
                endOffset,
                true));
        }

        private static CitationBounds TrimCitationBounds(string text, int startOffset, int endOffset)
        {
            int start = startOffset;
            int end = endOffset;

            while (start < end && (char.IsWhiteSpace(text[start]) || IsOpeningCitationWrapper(text[start])))
            {
                start += 1;
            }

            while (end > start && (char.IsWhiteSpace(text[end - 1]) || IsClosingCitationWrapper(text[end - 1])))
            {
                end -= 1;
            }

            return new CitationBounds(start, end);
        }

        private static CitationSegment GetCitationBacklinkSourceSegment(CitationSegment segment)
        {
            if (!segment.BacklinkStartOffset.HasValue || !segment.BacklinkEndOffset.HasValue)
            {
                return segment;
            }

            return new CitationSegment(segment.Item, segment.BacklinkStartOffset.Value, segment.BacklinkEndOffset.Value, true);
        }

        private static Word.Range RangeFromOffsets(Word.Document document, Word.Range sourceRange, int startOffset, int endOffset)
        {
            if (sourceRange == null || startOffset < 0 || endOffset <= startOffset)
            {
                return null;
            }

            int start = sourceRange.Start + startOffset;
            int end = sourceRange.Start + endOffset;
            if (start < sourceRange.Start || end > sourceRange.End || start >= end)
            {
                return null;
            }

            return document.Range(start, end);
        }

        private static void AddOrReplaceBookmark(Word.Document document, string bookmarkName, Word.Range range)
        {
            if (document.Bookmarks.Exists(bookmarkName))
            {
                document.Bookmarks[bookmarkName].Delete();
            }

            document.Bookmarks.Add(bookmarkName, range);
        }

        private static string BuildTemporaryCitationBookmarkName(int index)
        {
            return "Cite_Temp_" + Guid.NewGuid().ToString("N").Substring(0, 20) + "_" + index.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }

        private static Word.Hyperlink AddInternalHyperlink(
            Word.Document document,
            Word.Range range,
            string subAddress,
            string screenTip,
            bool preserveDisplayText = false)
        {
            object address = string.Empty;
            object sub = subAddress;
            object tip = screenTip;
            object textToDisplay = preserveDisplayText ? (object)(range.Text ?? string.Empty) : Type.Missing;
            object target = Type.Missing;
            return document.Hyperlinks.Add(range, ref address, ref sub, ref tip, ref textToDisplay, ref target);
        }

        private static bool RangeOverlapsAny(Word.Range range, IEnumerable<Word.Range> ranges)
        {
            if (range == null || ranges == null)
            {
                return false;
            }

            foreach (Word.Range candidate in ranges)
            {
                if (candidate == null)
                {
                    continue;
                }

                if (range.Start < candidate.End && range.End > candidate.Start)
                {
                    return true;
                }
            }

            return false;
        }

        private static IEnumerable<Word.Range> FindAllRanges(Word.Range sourceRange, string query, bool matchCase)
        {
            List<Word.Range> ranges = new List<Word.Range>();
            if (sourceRange == null || string.IsNullOrWhiteSpace(query))
            {
                return ranges;
            }

            Word.Range searchRange = sourceRange.Duplicate;
            int sourceEnd = sourceRange.End;
            while (searchRange.Start < sourceEnd)
            {
                searchRange.Find.ClearFormatting();
                bool found = searchRange.Find.Execute(
                    FindText: query,
                    MatchCase: matchCase,
                    MatchWholeWord: false,
                    MatchWildcards: false);

                if (!found || searchRange.End <= searchRange.Start)
                {
                    break;
                }

                ranges.Add(sourceRange.Document.Range(searchRange.Start, searchRange.End));
                int nextStart = searchRange.End;
                if (nextStart >= sourceEnd)
                {
                    break;
                }

                searchRange = sourceRange.Document.Range(nextStart, sourceEnd);
            }

            return ranges;
        }

        private static Word.Range FindFirstRange(Word.Range sourceRange, string query, bool matchCase)
        {
            return FindAllRanges(sourceRange, query, matchCase).FirstOrDefault();
        }

        private static bool ApplyCitationFormatting(Word.Range range, Color color, float fontSize)
        {
            if (range == null)
            {
                return false;
            }

            Word.Font font = range.Font;
            bool changed = false;
            changed |= SetFontColorIfNeeded(font, (Word.WdColor)ColorTranslator.ToOle(color));
            changed |= SetFontSizeIfNeeded(font, fontSize);
            changed |= SetFontUnderlineIfNeeded(font, Word.WdUnderline.wdUnderlineNone);
            return changed;
        }

        private static bool ApplyBibliographyBacklinkFormatting(Word.Range range)
        {
            if (range == null)
            {
                return false;
            }

            Word.Font font = range.Font;
            bool changed = false;
            changed |= SetFontColorIfNeeded(font, Word.WdColor.wdColorBlack);
            changed |= SetFontColorIndexIfNeeded(font, Word.WdColorIndex.wdBlack);
            changed |= SetFontUnderlineIfNeeded(font, Word.WdUnderline.wdUnderlineNone);
            return changed;
        }

        private static void EnsureZoteroHyperlinkStyles(Word.Document document, Color citationColor)
        {
            if (document == null)
            {
                return;
            }

            ApplyHyperlinkStyleFormatting(document, Word.WdBuiltinStyle.wdStyleHyperlink, citationColor);
            ApplyHyperlinkStyleFormatting(document, Word.WdBuiltinStyle.wdStyleHyperlinkFollowed, citationColor);
            ApplyNamedHyperlinkStyleFormatting(document, "Hyperlink", citationColor);
            ApplyNamedHyperlinkStyleFormatting(document, "FollowedHyperlink", citationColor);
            ApplyNamedHyperlinkStyleFormatting(document, "Followed Hyperlink", citationColor);
            ApplyNamedHyperlinkStyleFormatting(document, "\u8D85\u94FE\u63A5", citationColor);
            ApplyNamedHyperlinkStyleFormatting(document, "\u5DF2\u8BBF\u95EE\u7684\u8D85\u94FE\u63A5", citationColor);
            ApplyNamedHyperlinkStyleFormatting(document, "\u8BBF\u95EE\u8FC7\u7684\u8D85\u94FE\u63A5", citationColor);
        }

        private static void ApplyHyperlinkStyleFormatting(Word.Document document, Word.WdBuiltinStyle styleId, Color citationColor)
        {
            try
            {
                Word.Style style = document.Styles[styleId];
                ApplyHyperlinkStyleFormatting(style, citationColor);
            }
            catch
            {
            }
        }

        private static void ApplyNamedHyperlinkStyleFormatting(Word.Document document, string styleName, Color citationColor)
        {
            try
            {
                Word.Style style = document.Styles[styleName];
                ApplyHyperlinkStyleFormatting(style, citationColor);
            }
            catch
            {
            }
        }

        private static void ApplyHyperlinkStyleFormatting(Word.Style style, Color citationColor)
        {
            if (style == null)
            {
                return;
            }

            try
            {
                Word.Font font = style.Font;
                SetFontColorIfNeeded(font, (Word.WdColor)ColorTranslator.ToOle(citationColor));
                SetFontUnderlineIfNeeded(font, Word.WdUnderline.wdUnderlineNone);
            }
            catch
            {
            }
        }

        private static bool ResetZoteroFieldFormatting(Word.Range range)
        {
            if (range == null)
            {
                return false;
            }

            Word.Font font = range.Font;
            bool changed = false;
            changed |= SetFontColorIndexIfNeeded(font, Word.WdColorIndex.wdAuto);
            changed |= SetFontUnderlineIfNeeded(font, Word.WdUnderline.wdUnderlineNone);
            return changed;
        }

        private static bool SetFontColorIfNeeded(Word.Font font, Word.WdColor color)
        {
            if (font == null || IsFontColor(font, color))
            {
                return false;
            }

            font.Color = color;
            return true;
        }

        private static bool SetFontColorIndexIfNeeded(Word.Font font, Word.WdColorIndex colorIndex)
        {
            if (font == null || IsFontColorIndex(font, colorIndex))
            {
                return false;
            }

            font.ColorIndex = colorIndex;
            return true;
        }

        private static bool SetFontUnderlineIfNeeded(Word.Font font, Word.WdUnderline underline)
        {
            if (font == null || IsFontUnderline(font, underline))
            {
                return false;
            }

            font.Underline = underline;
            return true;
        }

        private static bool SetFontSizeIfNeeded(Word.Font font, float fontSize)
        {
            if (font == null || IsFontSize(font, fontSize))
            {
                return false;
            }

            font.Size = fontSize;
            return true;
        }

        private static bool IsFontColor(Word.Font font, Word.WdColor color)
        {
            try
            {
                return (int)font.Color == (int)color;
            }
            catch
            {
                return false;
            }
        }

        private static bool IsFontColorIndex(Word.Font font, Word.WdColorIndex colorIndex)
        {
            try
            {
                return font.ColorIndex == colorIndex;
            }
            catch
            {
                return false;
            }
        }

        private static bool IsFontUnderline(Word.Font font, Word.WdUnderline underline)
        {
            try
            {
                return font.Underline == underline;
            }
            catch
            {
                return false;
            }
        }

        private static bool IsFontSize(Word.Font font, float fontSize)
        {
            try
            {
                return Math.Abs(font.Size - fontSize) < 0.05f;
            }
            catch
            {
                return false;
            }
        }

        private static string SegmentTooltip(CitationItem item, string fallback)
        {
            if (!string.IsNullOrWhiteSpace(item.Title))
            {
                return item.Title;
            }

            return !string.IsNullOrWhiteSpace(item.Author) ? item.Author : fallback;
        }

        private static List<CitationItem> ParseCitationItems(string code)
        {
            string jsonText = CleanZoteroCitationJsonText(code);
            string itemsBlock = ExtractBalancedArray(jsonText, "\"citationItems\"");
            if (string.IsNullOrEmpty(itemsBlock))
            {
                return new List<CitationItem>();
            }

            List<CitationItem> items = new List<CitationItem>();
            foreach (string itemBlock in SplitTopLevelObjects(itemsBlock))
            {
                string itemDataBlock = ExtractBalancedObject(itemBlock, "\"itemData\"");
                string sourceBlock = string.IsNullOrEmpty(itemDataBlock) ? itemBlock : itemDataBlock;
                string title = StripHtml(ExtractJsonString(sourceBlock, "title"));
                CitationItem item = new CitationItem
                {
                    Id = ExtractJsonNumberString(sourceBlock, "id"),
                    Title = title,
                    Author = ExtractFirstAuthorFamily(sourceBlock),
                    Year = ExtractFirstIssuedYear(sourceBlock)
                };

                if (string.IsNullOrEmpty(item.Id))
                {
                    item.Id = ExtractJsonNumberString(itemBlock, "id");
                }

                if (!string.IsNullOrWhiteSpace(item.Id) ||
                    !string.IsNullOrWhiteSpace(item.Title) ||
                    !string.IsNullOrWhiteSpace(item.Author) ||
                    !string.IsNullOrWhiteSpace(item.Year))
                {
                    items.Add(item);
                }
            }

            return items;
        }

        private static string CleanZoteroCitationJsonText(string code)
        {
            string value = Regex.Replace(
                code ?? string.Empty,
                "^.*?ADDIN\\s+ZOTERO_ITEM\\s+CSL_CITATION\\s+",
                string.Empty,
                RegexOptions.IgnoreCase | RegexOptions.Singleline);
            value = Regex.Replace(value, "[\u0013\u0014][\\s\\S]*?\u0015", string.Empty);
            value = Regex.Replace(value, "[\u0000-\u0008\u000b\u000c\u000e-\u001f]", string.Empty);
            return value.Trim();
        }

        private static string ExtractBalancedArray(string text, string key)
        {
            int keyIndex = (text ?? string.Empty).IndexOf(key, StringComparison.Ordinal);
            if (keyIndex < 0)
            {
                return null;
            }

            int start = text.IndexOf("[", keyIndex + key.Length, StringComparison.Ordinal);
            return start < 0 ? null : ExtractBalancedText(text, start, '[', ']');
        }

        private static string ExtractBalancedObject(string text, string key)
        {
            int keyIndex = (text ?? string.Empty).IndexOf(key, StringComparison.Ordinal);
            if (keyIndex < 0)
            {
                return null;
            }

            int start = text.IndexOf("{", keyIndex + key.Length, StringComparison.Ordinal);
            return start < 0 ? null : ExtractBalancedText(text, start, '{', '}');
        }

        private static string ExtractBalancedText(string text, int start, char open, char close)
        {
            int depth = 0;
            bool inString = false;
            bool escaped = false;
            for (int index = start; index < text.Length; index += 1)
            {
                char character = text[index];
                if (inString)
                {
                    if (escaped)
                    {
                        escaped = false;
                    }
                    else if (character == '\\')
                    {
                        escaped = true;
                    }
                    else if (character == '"')
                    {
                        inString = false;
                    }

                    continue;
                }

                if (character == '"')
                {
                    inString = true;
                }
                else if (character == open)
                {
                    depth += 1;
                }
                else if (character == close)
                {
                    depth -= 1;
                    if (depth == 0)
                    {
                        return text.Substring(start, index - start + 1);
                    }
                }
            }

            return null;
        }

        private static List<string> SplitTopLevelObjects(string arrayText)
        {
            List<string> objects = new List<string>();
            int objectStart = -1;
            int depth = 0;
            bool inString = false;
            bool escaped = false;

            for (int index = 1; index < arrayText.Length - 1; index += 1)
            {
                char character = arrayText[index];
                if (inString)
                {
                    if (escaped)
                    {
                        escaped = false;
                    }
                    else if (character == '\\')
                    {
                        escaped = true;
                    }
                    else if (character == '"')
                    {
                        inString = false;
                    }

                    continue;
                }

                if (character == '"')
                {
                    inString = true;
                }
                else if (character == '{')
                {
                    if (depth == 0)
                    {
                        objectStart = index;
                    }

                    depth += 1;
                }
                else if (character == '}')
                {
                    depth -= 1;
                    if (depth == 0 && objectStart >= 0)
                    {
                        objects.Add(arrayText.Substring(objectStart, index - objectStart + 1));
                        objectStart = -1;
                    }
                }
            }

            return objects;
        }

        private static string ExtractJsonString(string text, string key)
        {
            Match match = Regex.Match(
                text ?? string.Empty,
                "\"" + Regex.Escape(key) + "\"\\s*:\\s*\"((?:\\\\.|[^\"\\\\])*)\"");
            return match.Success ? UnescapeJsonString(match.Groups[1].Value) : string.Empty;
        }

        private static string ExtractJsonNumberString(string text, string key)
        {
            Match match = Regex.Match(
                text ?? string.Empty,
                "\"" + Regex.Escape(key) + "\"\\s*:\\s*\"?([0-9]+)\"?");
            return match.Success ? match.Groups[1].Value : string.Empty;
        }

        private static string ExtractFirstAuthorFamily(string text)
        {
            string authorBlock = ExtractBalancedArray(text, "\"author\"");
            return string.IsNullOrEmpty(authorBlock) ? string.Empty : ExtractJsonString(authorBlock, "family");
        }

        private static string ExtractFirstIssuedYear(string text)
        {
            string issuedBlock = ExtractBalancedObject(text, "\"issued\"");
            if (string.IsNullOrEmpty(issuedBlock))
            {
                return string.Empty;
            }

            Match match = Regex.Match(issuedBlock, "\"date-parts\"\\s*:\\s*\\[\\s*\\[\\s*\"?([0-9]{4})\"?");
            return match.Success ? match.Groups[1].Value : string.Empty;
        }

        private static string UnescapeJsonString(string value)
        {
            string result = value
                .Replace("\\\"", "\"")
                .Replace("\\\\", "\\")
                .Replace("\\/", "/")
                .Replace("\\b", "\b")
                .Replace("\\f", "\f")
                .Replace("\\n", "\n")
                .Replace("\\r", "\r")
                .Replace("\\t", "\t");

            return Regex.Replace(result, "\\\\u([0-9a-fA-F]{4})", match =>
                ((char)Convert.ToInt32(match.Groups[1].Value, 16)).ToString());
        }

        private static string BuildBookmarkName(CitationItem item)
        {
            string source = string.Join(
                "|",
                item.Id ?? string.Empty,
                item.Title ?? string.Empty,
                item.Year ?? string.Empty,
                item.Author ?? string.Empty);
            string id = CleanBookmarkPart(item.Id ?? "NoId", 10);
            string title = CleanBookmarkPart(item.Title ?? "NoTitle", 14);
            string year = CleanBookmarkPart(item.Year ?? "NoYear", 6);
            string author = CleanBookmarkPart(item.Author ?? "NoAuthor", 10);
            return TruncateBookmarkNameWithHash("Cite_" + id + "_" + title + "_" + year + "_" + author, source);
        }

        private static string BuildCitationBookmarkName(CitationItem item)
        {
            string baseName = BuildBookmarkName(item);
            if (baseName.StartsWith("Cite_", StringComparison.OrdinalIgnoreCase))
            {
                baseName = baseName.Substring("Cite_".Length);
            }

            return TruncateBookmarkNameWithHash("Cite_Back_" + baseName, baseName);
        }

        private static string CleanBookmarkPart(string value, int maxLength)
        {
            string cleaned = Regex.Replace(value ?? string.Empty, "[^A-Za-z0-9_]", "_");
            if (Regex.IsMatch(cleaned, "^[0-9]"))
            {
                cleaned = "_" + cleaned;
            }

            if (string.IsNullOrEmpty(cleaned))
            {
                cleaned = "_";
            }

            return cleaned.Length <= maxLength ? cleaned : cleaned.Substring(0, maxLength);
        }

        private static string TruncateBookmarkNameWithHash(string value, string source)
        {
            const int maxLength = 40;
            string hashSuffix = "_" + SimpleHash(source);
            if (value.Length <= maxLength - hashSuffix.Length)
            {
                return value + hashSuffix;
            }

            return value.Substring(0, maxLength - hashSuffix.Length) + hashSuffix;
        }

        private static string SimpleHash(string value)
        {
            unchecked
            {
                uint hashValue = 2166136261;
                string source = value ?? string.Empty;
                for (int index = 0; index < source.Length; index += 1)
                {
                    hashValue ^= source[index];
                    hashValue *= 16777619;
                }

                return hashValue.ToString("x8", System.Globalization.CultureInfo.InvariantCulture);
            }
        }

        private static string StripHtml(string value)
        {
            return Regex.Replace(value ?? string.Empty, "</?(i|sub|sup|span|b|strong|em)[^>]*>", string.Empty, RegexOptions.IgnoreCase).Trim();
        }

        private static string ShortenSearchText(string value, int maxLength)
        {
            string compacted = Regex.Replace(value ?? string.Empty, "\\s+", " ").Trim();
            if (compacted.Length <= maxLength)
            {
                return compacted;
            }

            string shortened = compacted.Substring(0, maxLength);
            int lastSpace = shortened.LastIndexOf(" ", StringComparison.Ordinal);
            return (lastSpace > 10 ? shortened.Substring(0, lastSpace) : shortened).Trim();
        }

        private static string NormalizeCitationDisplayText(string text)
        {
            return Regex.Replace(text ?? string.Empty, "\\s+", " ").Trim();
        }

        private static string NormalizeBibliographyMatchText(string text)
        {
            return Regex.Replace(text ?? string.Empty, "\\s+", " ").Trim().ToLowerInvariant();
        }

        private static bool ContainsBibliographyToken(string text, string token)
        {
            if (string.IsNullOrWhiteSpace(text) || string.IsNullOrWhiteSpace(token))
            {
                return false;
            }

            string escaped = Regex.Escape(token);
            return Regex.IsMatch(text, "(^|[^a-z0-9])" + escaped + "($|[^a-z0-9])", RegexOptions.IgnoreCase);
        }

        private static string NormalizeStyleId(string styleId)
        {
            if (string.IsNullOrWhiteSpace(styleId))
            {
                return string.Empty;
            }

            string value = styleId.Trim().ToLowerInvariant();
            if (value.StartsWith("china-national-standard-gb-t-7714-2015-numeric", StringComparison.Ordinal))
            {
                return "china-national-standard-gb-t-7714-2015-numeric";
            }

            if (value.StartsWith("china-national-standard-gb-t-7714-2015-author-date", StringComparison.Ordinal))
            {
                return "china-national-standard-gb-t-7714-2015-author-date";
            }

            if (value.StartsWith("optics-express", StringComparison.Ordinal) ||
                value.Contains("the-optical-society"))
            {
                return "optics-express";
            }

            if (value.StartsWith("opto-electronic-advances", StringComparison.Ordinal))
            {
                return "opto-electronic-advances";
            }

            return value;
        }

        private static string LastPathSegment(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            string[] parts = value.Split('/');
            return parts.Length == 0 ? value : parts[parts.Length - 1];
        }

        private static CitationItem GetItemAt(List<CitationItem> items, int index)
        {
            return index >= 0 && index < items.Count ? items[index] : null;
        }

        private static string SafeFieldCode(Word.Field field)
        {
            try
            {
                return field.Code.Text ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static bool IsZoteroCitationCode(string code)
        {
            return Regex.IsMatch(code ?? string.Empty, "ADDIN\\s+ZOTERO_ITEM", RegexOptions.IgnoreCase);
        }

        private static bool IsZoteroBibliographyCode(string code)
        {
            return Regex.IsMatch(code ?? string.Empty, "ADDIN\\s+ZOTERO_BIBL", RegexOptions.IgnoreCase);
        }

        private static string NormalizeCiteName(string value)
        {
            return (value ?? string.Empty).TrimStart('#');
        }

        private static bool IsCiteName(string value)
        {
            return NormalizeCiteName(value).StartsWith("Cite_", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsCitationBackBookmarkName(string value)
        {
            return NormalizeCiteName(value).StartsWith("Cite_Back_", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsLegacyMacroBookmarkName(string value)
        {
            string normalized = NormalizeCiteName(value);
            return Regex.IsMatch(normalized ?? string.Empty, "^[A-Za-z_][A-Za-z0-9_]{0,43}$");
        }

        private static bool IsCitationSeparator(char value)
        {
            return value == ',' || value == '\uff0c' || value == ';' || value == '\uff1b';
        }

        private static bool IsCommaLikeChar(char value)
        {
            return value == ',' || value == '\uff0c';
        }

        private static bool IsSemicolonLikeChar(char value)
        {
            return value == ';' || value == '\uff1b';
        }

        private static bool IsRangeDash(char value)
        {
            return value == '-' ||
                value == '\u2010' ||
                value == '\u2011' ||
                value == '\u2012' ||
                value == '\u2013' ||
                value == '\u2014' ||
                value == '\u2212' ||
                value == '\ufe63' ||
                value == '\uff0d';
        }

        private static bool IsOpeningCitationWrapper(char value)
        {
            return value == '[' || value == '\uff3b' || value == '(' || value == '{' || value == '\uff08' || value == '\u3014';
        }

        private static bool IsClosingCitationWrapper(char value)
        {
            return value == ']' || value == '\uff3d' || value == ')' || value == '}' || value == '\uff09' || value == '\u3015';
        }

        internal static Color DefaultColor
        {
            get { return Color.Red; }
        }

        internal static float DefaultFontSize
        {
            get { return DefaultCitationFontSize; }
        }
    }

    internal sealed class LinkResult
    {
        internal int Linked { get; set; }
        internal int LinkedBacklinks { get; set; }
        internal int SkippedMultiItem { get; set; }
        internal int SkippedMissingBibliography { get; set; }
        internal int FailedBibliographyMatch { get; set; }
        internal int FailedCitationRange { get; set; }
        internal int SkippedCompressedItems { get; set; }
    }

    internal sealed class RemoveResult
    {
        internal int LinksRemoved { get; set; }
        internal int BookmarksRemoved { get; set; }
        internal int Recolored { get; set; }
    }

    internal sealed class ZoteroFields
    {
        internal ZoteroFields()
        {
            Citations = new List<ZoteroFieldInfo>();
        }

        internal List<ZoteroFieldInfo> Citations { get; private set; }
        internal Word.Field Bibliography { get; set; }
    }

    internal sealed class ZoteroFieldInfo
    {
        internal Word.Field Field { get; set; }
        internal string Code { get; set; }
        internal string Text { get; set; }
        internal List<CitationItem> CitationItems { get; set; }
    }

    internal sealed class CitationItem
    {
        internal string Id { get; set; }
        internal string Title { get; set; }
        internal string Author { get; set; }
        internal string Year { get; set; }
    }

    internal sealed class BibliographyEntryMatch
    {
        internal BibliographyEntryMatch(int startOffset, int endOffset, string text)
        {
            StartOffset = startOffset;
            EndOffset = endOffset;
            Text = text;
        }

        internal int StartOffset { get; private set; }
        internal int EndOffset { get; private set; }
        internal string Text { get; private set; }
        internal int Score { get; set; }
    }

    internal sealed class CitationItemReferenceComparer : IEqualityComparer<CitationItem>
    {
        public bool Equals(CitationItem x, CitationItem y)
        {
            return ReferenceEquals(x, y);
        }

        public int GetHashCode(CitationItem obj)
        {
            return obj == null ? 0 : System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(obj);
        }
    }

    internal sealed class CitationSegment
    {
        internal CitationSegment(CitationItem item, int startOffset, int endOffset, bool visible)
        {
            Item = item;
            StartOffset = startOffset;
            EndOffset = endOffset;
            Visible = visible;
        }

        internal CitationItem Item { get; private set; }
        internal int StartOffset { get; private set; }
        internal int EndOffset { get; private set; }
        internal bool Visible { get; private set; }
        internal int? BacklinkStartOffset { get; set; }
        internal int? BacklinkEndOffset { get; set; }
    }

    internal sealed class PendingCitationLink
    {
        internal PendingCitationLink(string temporaryBookmarkName, string bibliographyBookmarkName, string screenTip)
        {
            TemporaryBookmarkName = temporaryBookmarkName;
            BibliographyBookmarkName = bibliographyBookmarkName;
            ScreenTip = screenTip;
        }

        internal string TemporaryBookmarkName { get; private set; }
        internal string BibliographyBookmarkName { get; private set; }
        internal string ScreenTip { get; private set; }
    }

    internal sealed class PendingBibliographyBacklink
    {
        internal PendingBibliographyBacklink(string bibliographyBookmarkName, string citationBookmarkName)
        {
            BibliographyBookmarkName = bibliographyBookmarkName;
            CitationBookmarkName = citationBookmarkName;
        }

        internal string BibliographyBookmarkName { get; private set; }
        internal string CitationBookmarkName { get; private set; }
    }

    internal sealed class PendingBibliographyBacklinkRange
    {
        internal PendingBibliographyBacklinkRange(int start, int end, string citationBookmarkName)
        {
            Start = start;
            End = end;
            CitationBookmarkName = citationBookmarkName;
        }

        internal int Start { get; private set; }
        internal int End { get; private set; }
        internal string CitationBookmarkName { get; private set; }
    }

    internal sealed class NumericToken
    {
        internal int Value { get; set; }
        internal int StartOffset { get; set; }
        internal int EndOffset { get; set; }
        internal NumericSeparator SeparatorAfter { get; set; }
    }

    internal sealed class RangeStart
    {
        internal RangeStart(int itemIndex, NumericToken token)
        {
            ItemIndex = itemIndex;
            Token = token;
        }

        internal int ItemIndex { get; private set; }
        internal NumericToken Token { get; private set; }
    }

    internal sealed class CitationBounds
    {
        internal CitationBounds(int startOffset, int endOffset)
        {
            StartOffset = startOffset;
            EndOffset = endOffset;
        }

        internal int StartOffset { get; private set; }
        internal int EndOffset { get; private set; }
    }

    internal enum NumericSeparator
    {
        None,
        Comma,
        Dash,
        End
    }
}
