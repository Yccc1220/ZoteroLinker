# Zotero Linker

[中文](README.md) | English

Zotero Linker is a Windows suite for Zotero and Office. It includes a Word citation-navigation add-in and a PowerPoint citation add-in. The Word add-in links numeric, APA, MLA, Chicago, and other in-text citation formats with bibliography entries and can add DOI hyperlinks manually, while the PowerPoint add-in brings Zotero style selection, item selection, numbered citations, and bibliography generation into slide decks.

The product family uses `zoterolinker`; the Word and PowerPoint products use `zoterolinkerword` and `zoterolinkerppt`. Existing internal VSTO identifiers remain unchanged for installed-version compatibility.

## The Gaps It Solves

### Word: citations exist, but navigation is awkward

Zotero can insert citations and generate bibliographies in Word, but numeric, APA, MLA, and Chicago citations do not always provide convenient two-way navigation to their bibliography entries. The add-in handles forms such as `[1]`, `[2-4]`, `(Smith, 2020)`, `Smith (2020)`, and `(Smith 23; Jones 41)`, reducing repeated scrolling and searching in long documents.

The Word add-in preserves Zotero fields and existing text formatting, creates links in both directions, and provides DOI linking, explicit formatting controls, and link cleanup.

### PowerPoint: no equivalent Zotero citation workflow

Zotero's official Office integration is centered on Word. PowerPoint does not provide the same style-selection, item-search, multi-citation, automatic-numbering, and bibliography-refresh workflow. Users often type `[1]` manually and copy bibliography text, which makes numbering and style consistency fragile when references change.

The PowerPoint add-in uses Zotero's local citing protocol to open Zotero's document preferences and item picker, insert single or multiple citations into slides, and generate or refresh a bibliography slide. It fills the integration gap between Zotero and presentations without replacing Zotero.

## Word Add-in

- Link in-text citations to bibliography entries and back again.
- Support numeric formats such as `[1]`, `[1,3,5]`, and `[2-4]`, plus APA, MLA, Chicago author-date, narrative, and multi-source citations.
- Use Zotero field data to handle visible and hidden items in compressed citations.
- Manually add hyperlinks for `doi:`, `https://doi.org/`, and bare DOI values in bibliography entries.
- Apply citation color and font size separately from Options; link operations do not change formatting automatically.
- Remove generated links and bookmarks while preserving Zotero fields, colors, underlines, and font sizes.
- Support Microsoft Office Word and WPS Word/Writer.

## PowerPoint Add-in

- Open Zotero document preferences on first use to select a CSL style.
- Open Zotero's item picker with single- and multi-select support.
- Insert numbered citations such as `[1]` or `[1–4]` into the active slide.
- Create or update a bibliography on a `References` slide.
- Support `Document Preferences` and `Refresh` commands.
- Persist Zotero document and field metadata inside the presentation for later refreshes.

## Compatibility

| Add-in | Supported environment |
| --- | --- |
| Word | Windows, Microsoft Office Word, WPS Word/Writer, Zotero |
| PowerPoint | Windows, Microsoft Office PowerPoint, Zotero |
| WPS Presentation | The installer writes the WPS Presentation (`WPP`) add-in whitelist; actual loading depends on the VSTO compatibility of the installed WPS version |

The PowerPoint add-in requires Zotero to be running with local application communication enabled. Citations and bibliography entries are stored as PowerPoint text shapes rather than Word fields.

## Installation

Download the required installer from [GitHub Releases](https://github.com/Yccc1220/ZoteroLinker/releases/latest):

```text
ZoteroLinkerWordSetup.exe  # Word / WPS Writer
ZoteroLinkerPptSetup.exe   # PowerPoint / WPS Presentation compatibility registration
```

Run the installer as administrator, then reopen the relevant Office or WPS application.

## PowerPoint Usage

1. Start Zotero and open a PowerPoint presentation.
2. Select `Insert Citation` on the `Zotero Linker` ribbon tab.
3. Choose a citation style on first use, then select one or more items in Zotero.
4. Select `Add Bibliography` to create the references slide.
5. Select `Refresh` after citations or document preferences change.

## Repository Contents

- `Zotero-linker-word/`: Word VSTO add-in source and installer.
- `Zotero-linker-ppt/`: PowerPoint VSTO add-in source and installer.
- `release/`: Word and PowerPoint installers with SHA-256 checksums.
- `.github/workflows/release.yml`: GitHub Release workflow.
- `site/`: GitHub Pages download page.

