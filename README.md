# MFSLB — Make Finance Secretary's Life Better

MFSLB is a Google Apps Script automation built for the University of Toronto Engineering Society Finance Committee. It automates repetitive administrative work around monthly funding applications, document organization, reviewer assignment, meeting preparation, and post-meeting budget tracking.

The project was developed by Finance Secretary Suyeon Park with Kenneth Sulimro during 2023–2024. The workflow was created to replace a manual preparation process that previously took several hours with a mostly automated sequence of operations.

## What It Automates

1. **Monthly form management** — archives the previous response sheet, resets the form, and prepares the next monthly intake cycle.
2. **Application file organization** — categorizes uploaded files and moves them into application-type folders.
3. **Data extraction** — reads form responses and transforms the relevant fields into structured application data.
4. **Reviewer assignment** — assigns pairs of Finance Committee members to applications for review.
5. **Meeting-document generation** — creates meeting agenda/minutes documents from templates and application data.
6. **Comment-sheet generation** — prepares a review sheet populated with applications and assigned reviewers.
7. **Budget tracking** — records approved/requested funding information after meetings and updates aggregate formulas.

## Workflow

```text
Google Form submission
        |
        v
Google Apps Script
   |        |        |
   v        v        v
Drive     Sheets    Docs
files     review    meeting
organized sheets    agenda/minutes
        \    |    /
             v
       Budget tracker
```

## Configuration

The source code intentionally does not hard-code the main Google Form, Drive, Sheet, or document-template resource IDs. They are loaded from Google Apps Script **Script Properties**.

The deployment expects properties including:

```text
FORM_ID
INTERNAL_DRIVE_ID
PARENT_FOLDER_ID
MINUTES_TEMPLATE_ID
COMMENT_SHEET_ID
TRACKER_SHEET_ID
```

Runtime-created values such as the current month's folder ID and URL are also stored through `PropertiesService`.

Production Google resource IDs, committee documents, form responses, and other private Finance Committee data are not intended to be committed to this repository.

## Source Layout

- `main.gs` — primary Finance Committee workflow and document automation.
- supporting Apps Script files — helper/configuration logic used by the workflow.

## Demo Videos

- [Finance Committee Workflow](https://www.youtube.com/watch?v=0HXMRhMQp8E)
- [MFSLB Project](https://www.youtube.com/watch?v=mau6UkN7GC0)

## Access

The production Apps Script project and its associated Google Workspace resources are restricted to authorized Engineering Society accounts. This public repository is intended to demonstrate the automation logic without exposing access to the underlying committee data.

## Notes

MFSLB was developed around a real committee workflow whose forms and spreadsheet layouts can change over time. Column mappings and template assumptions therefore need to be reviewed when the underlying Finance Committee documents change.
