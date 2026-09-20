# Judicial State Government Organization Matching & Audit

A Python-based data matching and validation workflow for identifying and verifying organization records across SQL Server database tables using organization names and state information.

## Overview

This project automates organization matching between an Excel input file and two SQL Server organization data sources.

For each organization in the input Excel file, the workflow generates search keywords, retrieves database candidates, validates candidate names, assigns a confidence level, and writes the results back to the same Excel workbook.

The database workflow is read-only. The script does not insert, update, or delete database records.

## How It Works

The matching process follows these steps:

1. Read organization `Name` and `State` from an Excel file.
2. Generate candidate search keywords from the organization name.
3. Prioritize distinctive words, including words appearing after a `" - "` separator.
4. Remove common noise words and state names/abbreviations from search candidates.
5. Truncate long keywords into search stems for SQL `LIKE` searches.
6. Search one keyword at a time against the database.
7. Calculate a name-match score for returned candidates.
8. Validate candidates using exact or prefix-based token matching.
9. Continue through additional keywords when a candidate does not verify.
10. Assign a confidence level to verified matches.
11. Flag multiple verified candidates instead of automatically selecting one.
12. Write the results back to the Excel workbook.

## Matching Logic

### Keyword Generation

The script creates an ordered list of candidate keywords from each organization name.

The process:

- Removes common filler words.
- Removes the state's name and abbreviation.
- Prioritizes words appearing after `" - "` because they may represent distinctive branches, counties, or divisions.
- Prioritizes longer words.
- Removes duplicate keywords.
- Truncates keywords to a configurable search length.

For example, a word such as:

```text
Judicial
