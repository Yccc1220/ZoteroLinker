# Zotero Linker

[中文](README.md) | English

Zotero Linker is a Windows Word add-in that improves citation navigation in Zotero-based documents. It is designed for academic writing, thesis preparation, literature reviews, grant proposals, and technical reports with many references. The add-in has been adapted for Microsoft Office Word and WPS Word/Writer on Windows.

## The Gap It Solves

Zotero is excellent at managing references, inserting citations, and generating bibliographies. However, in Word documents, Zotero citations are not always convenient to navigate. Readers often need to jump from an in-text citation such as `[1]`, `[2-4]`, or `[3, 5, 8]` to the corresponding bibliography entry, and sometimes back from a reference entry to where it is cited in the main text.

In long documents, manually scrolling between the main text and the bibliography becomes slow and distracting. The more references a document contains, the more this affects writing, reviewing, and editing efficiency.

Zotero Linker does not replace Zotero. Instead, it fills this navigation and formatting gap by turning Zotero-generated Word documents into documents that are easier to read, review, and maintain.

## Key Features

- Link in-text citations to bibliography entries.
- Link bibliography entries back to the corresponding in-text citation.
- Support common numeric citation formats such as `[1]`, `[1,3,5]`, and `[2-4]`.
- Handle compressed numeric citations by using Zotero field information.
- Repair citation formatting, including color, underline, and font size.
- Remove generated links and bookmarks while preserving original Zotero citation fields.
- Customize citation color and font size.
- Show operation status directly in the Word ribbon, including linked items, backlinks, hidden items, and failed matches.

## Compatibility

- Operating system: Windows
- Supported editors: Microsoft Office Word, WPS Word/Writer
- Reference manager: Zotero
- Recommended document type: Word documents that still contain Zotero citation fields and bibliography fields

Note: If Zotero citations have been converted to plain text, the add-in may not be able to read the full Zotero field data, and some advanced features may not work.

## Installation

Download the installer from GitHub Releases:

```text
ZoteroLinkerSetup-1.0.0.exe
```

Run the installer as administrator, then reopen Microsoft Word or WPS Word/Writer.

## Basic Usage

1. Install the add-in and open Microsoft Word or WPS Word/Writer.
2. Open a document that contains Zotero citations and a Zotero bibliography.
3. Find the `Zotero Linker` tab in the ribbon.
4. Click `Link Citations` to create links between in-text citations and bibliography entries.
5. Use `Ctrl + Click` to navigate between citations and references.
6. Use `Remove Links` to remove generated links when needed.
7. Use `Repair Formatting` to restore citation color, underline, and font size.
8. Use `Options` to customize citation color and font size.

## Why It Matters

For short documents, manually checking references may be acceptable. For long papers, theses, and reviews with dozens or hundreds of references, citation navigation becomes a real productivity issue. Zotero Linker helps authors, reviewers, supervisors, and collaborators move between citations and references more efficiently.

In short:

> Zotero creates correct citations. Zotero Linker makes those citations easier to navigate, review, and maintain in Word documents.

## Repository Contents

- `release/`: built Windows installer and SHA-256 checksum
- `Zotero-linker/`: VSTO add-in source code
- `README.md`: Chinese documentation
- `README.en.md`: English documentation

## Notes

This add-in is mainly intended for Windows desktop writing workflows where Zotero is used together with Word or WPS Word/Writer.


