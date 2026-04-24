📧 Project Email Exporter (Graph-Based)

Version: v1.4.3

🧭 Overview

The Project Email Exporter is a PowerShell-based tool (with GUI) designed to:

Export project-related email conversations from Microsoft 365 mailboxes
Organise them into structured folders by Client → Project → Threads
Move completed project folders into a PREDELETE archive location
Maintain audit logs and CSV tracking for validation and recovery

It uses Microsoft Graph API for mailbox access and supports delegated access scenarios (e.g. service accounts with shared mailbox permissions).

🚀 Key Features
📥 Export Emails by Project
Reads a CSV input file (ProjectsToArchive.csv)
Matches mailbox folders based on project naming
Exports emails grouped by conversation thread
Outputs structured folders per project
📂 Folder-Based Archiving (v1.3+)
Moves entire matched mail folders (not loose emails)

Preserves structure inside:

Inbox\PREDELETE\<Project Folder>
🧾 Logging & Audit Tracking (v1.4+)

Automatically generates logs per run:

<OutputFolder>\Logs\<RunID>\
Includes:
Run summary CSV
Per-project CSV logs
Event/log transcript file
📊 CSV Status Tracking

Creates an updated CSV:

ProjectsToArchive_UPDATED_<RunID>.csv

Adds:

Category
Archiving Done
Review Required
Status/Error Info
Matched Folder
Processed At

✔ Updates after each project → safe against crashes mid-run

🧪 Dry Run Mode
Preview actions without making changes
Displays:
Matched folders
Email counts
Planned exports/moves
🔐 Microsoft Graph Integration
Uses:
Mail.ReadWrite
Mail.ReadWrite.Shared
Supports:
Delegated mailbox access
Cross-mailbox operations (with correct permissions)
📁 Folder Structure Output

Example:

C:\M365Export\
└── Client Name\
    └── 8868 – Project Name\
        ├── Thread 1\
        │   ├── Email1.eml
        │   ├── Email2.eml
        ├── Thread 2\
        └── ...

Logs:

C:\M365Export\Logs\2026-04-24-RunID\
├── RunSummary.csv
├── RunEvents.log
├── 8868_ProjectName.csv
🧾 CSV Input Format

Example:

Category,Project Number,Project Name,Client
ARCHIVING,8868,Speedway Wheel Analysis,Mag & Tyre Hastings
⚠️ Known Behaviours / Gotchas
🔸 Graph vs Exchange Permissions
Full Access in Exchange ≠ Graph Write Access

Must connect with:

Connect-MgGraph -Scopes "Mail.ReadWrite","Mail.ReadWrite.Shared","User.Read"
🔸 Special Characters in Names

Project names like:

[15x7]

previously broke logging/export

✔ Fixed in v1.4.x via:

Safe filename sanitisation
Removal of wildcard interpretation issues
🔸 Partial Export Handling
If a thread fails mid-export:
Logged as failed
CSV updated with "Review Required"
Script continues
🧪 Example Usage
GUI Mode
Launch script
Connect to Graph
Load CSV
Set output folder
Run:
Run Export
or Move Folder to PREDELETE
CLI Mode
pwsh.exe -ExecutionPolicy Bypass -File Export-ProjectEmails.ps1 `
  -MailboxUserId "clientData@domain.com" `
  -CsvPath "C:\ProjectsToArchive.csv" `
  -OutputFolder "C:\M365Export" `
  -MoveToPredelete
🧠 Design Philosophy
Fail-safe over fail-fast
Always produce usable output/logs even if incomplete
Designed for real-world messy data
Prioritises auditability and traceability
🛠️ Wishlist / Future Improvements
📊 Reporting Engine (High Value)

Add GUI button:

“Generate Report”

Aggregates logs into:
✔ Successfully exported
⚠ Skipped
❌ Failed
🔁 Partial

Output:

Report_<RunID>.html / .csv
📄 Full Transcript Logging
Capture everything shown in console UI

Save as:

FullTranscript.log
Useful for:
audits
troubleshooting
customer evidence
🧠 Smarter Matching Logic
Fuzzy matching for project folders
Detect:
duplicates
ambiguous names
Flag before processing
📁 Folder Validation Mode
Scan mailbox first
Show:
Missing folders
Multiple matches
Naming inconsistencies
🧾 Retry Engine
Retry failed messages:
X attempts
exponential backoff
Optional toggle
🖥️ GUI Improvements
Maximise / resizable window
Progress bars per project
Live log panel (scrolling)
Filterable project list
🔍 Pre-Delete Review Mode

Move to:

PREDELETE\Review\

Mark CSV as:

Pending Review
🔐 App Registration Mode
Support App-only auth
Remove dependency on interactive login
Ideal for automation / scheduled jobs
⏱️ Performance Improvements
Parallel processing for:
folder scans
exports
Graph batching where possible
🧩 BookStack / Documentation Export
Auto-generate:
HTML summary
upload-ready content
Useful for your current doc workflow
🧯 Troubleshooting
403 Access Denied
Missing Graph scopes
Not using Mail.ReadWrite.Shared
App consent not granted
Folder Not Found
Naming mismatch
Folder depth too low
Partial Exports
Usually Graph API limits / transient failures
Check logs per project
📌 Final Notes

This tool is built for practical mailbox cleanup + archival workflows, not just clean demo data.

It assumes:

inconsistent naming
partial permissions
real-world mess

…and handles it accordingly.
