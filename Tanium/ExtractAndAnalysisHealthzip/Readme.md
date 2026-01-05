# ExtractAndAnalyzeHealthZip

## Overview
**ExtractAndAnalyzeHealthZip** is a utility designed to:
- Extract a **Health ZIP** package provided by a client.
- Clean and normalize extracted files (remove noise, standardize structure/names).
- Search for specific **occurrences/patterns** inside log files to support troubleshooting and analysis.

## Features
- **Health ZIP extraction**
  - Unpacks the client Health ZIP into a structured workspace.
- **File cleanup**
  - Removes irrelevant or temporary files.
  - Optionally normalizes filenames and folders for consistent review.
- **Log analysis**
  - Searches for occurrences (keywords, strings, regex patterns) across log files.
  - Produces readable results to help identify issues quickly.