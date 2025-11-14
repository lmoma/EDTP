# LinkCheck – Word VBA Hyperlink Validator

## Overview
**LinkCheck** is a Word VBA macro that validates all hyperlinks in a Word document, including those in footnotes. It checks each unique URL only once, applies results to all instances, and generates a detailed report.

## Features
- **Deduplication**: Avoids redundant checks by validating unique URLs only.
- **HEAD-first optimization**: Uses HTTP HEAD requests for faster validation.
- **Color-coded results**:
  - ✅ Green: Valid link
  - ⚠️ Yellow: Failed check or error
  - ❌ Red: Invalid symbol detected
- **Context snippets**: Shows surrounding text for each link.
- **Performance tuning**:
  - Adjustable timeouts
  - Rate limiting between requests
  - Disables Word features during execution for speed
- **Abort option**: Press `ESC` to cancel validation mid-run.

## How It Works
1. Collects all hyperlinks from the main body and footnotes.
2. Normalizes URLs and removes duplicates.
3. Validates each unique URL using WinHTTP.
4. Generates a new Word document with:
   - A table of results (or plain text if table creation fails)
   - Display text, URL, context, and status.

## Requirements
- Microsoft Word (VBA-enabled)
- Windows (uses `WinHttp.WinHttpRequest` and `kernel32` API)
- Macro security settings allowing execution

## Installation
1. Open Word and press `Alt + F11` to open the VBA editor.
2. Insert a new module and paste the code from `LinkCheck.txt`.
3. Save and close the editor.

## Usage
- Open the document you want to check.
- Run `ValidateHyperlinks` from the Macros menu (`Alt + F8`).
- Wait for the process to complete. A new document will display the results.

## License
MIT License – Feel free to use, modify, and share.
