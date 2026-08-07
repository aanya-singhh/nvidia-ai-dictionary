# Weekly AI Report portable handoff

This folder contains the report generator, the current AI model watchlist, the
website, the training runbook, and the GitHub Pages publisher. Keep the folder
structure unchanged.

## What must be installed or signed in

1. **Codex or another tool-enabled research assistant**
   - Sign in to the tool used to research and curate the weekly newsletter.
   - It must be able to browse official company blogs, release notes, GitHub
     releases, documentation, and research papers.
   - The automated discovery pass is not a substitute for opening and verifying
     primary sources.

2. **Python 3.11 or newer**
   - On Windows, install Python and select **Add Python to PATH**.
   - From PowerShell in this folder, run:

     ```powershell
     python -m pip install -r requirements.txt
     ```

3. **Microsoft Outlook desktop**
   - Required only for automatic Outlook draft creation.
   - Sign in to a working Outlook profile and open Outlook once before the first
     production run.
   - New Outlook may not expose the required desktop COM automation. Use classic
     Outlook for Windows when draft creation fails.
   - The workflow saves a draft; it never sends the message automatically.

4. **Git and GitHub**
   - Install Git for Windows.
   - Sign in to GitHub through Git Credential Manager when the first push asks.
   - The account must have write access to the configured repository.
   - GitHub credentials or access tokens must never be stored in this folder.

5. **Microsoft Word**
   - Optional for generation, but recommended for reviewing the final DOCX.

No Microsoft Forms, Excel subscription workbook, or subscriber spreadsheet is
required by the production workflow.

## First-time setup

1. Copy `local_config.example.json` to `local_config.json`.
2. Edit `local_config.json`:
   - Replace the email recipient.
   - Replace the subscription URL.
   - Set the GitHub repository URL and branch.
   - Set `publish_website` to `true` only after the GitHub Page is configured.
3. Review the data-handling note below. Set
   `approve_public_model_search` to `true` only when authorized.
4. Install the Python requirements.
5. Follow `SETUP-GITHUB-PAGES.md` if this will publish a website.
6. Read the Word training document in `training`.

`local_config.json` is intentionally excluded from the distributed starter
configuration. It may contain organization-specific addresses and URLs; do not
email it after it has been personalized.

## Weekly operating sequence

1. Copy the newest `data/weekly_ai_report_input_YYYY-MM-DD.json` to a file named
   for the current Monday.
2. Update its inclusive date range, output filename, and subject.
3. Clear old report items and research the current week.
4. Populate the four required sections with verified sources:
   - NVIDIA AI News
   - Watchlist Updates
   - Technological Advancements in AI
   - New AI Models and Rumors
5. Run a safe QA pass:

   ```powershell
   powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Run-WeeklyAIReport.ps1 `
     -InputPath .\data\weekly_ai_report_input_YYYY-MM-DD.json `
     -SkipEmailDraft `
     -SkipWebsitePublish
   ```

6. Inspect the generated DOCX in `output`, the email fallback/preview if any,
   and `site/index.html`.
7. Run production without the two skip switches.
8. Review the Outlook draft, GitHub commit, and public site. Leave the email
   unsent until a person approves it.

## Public-search data handling

The discovery script sends each public `Name` and `Maker` value from
`data/ai-models.csv` to Bing RSS with release/update keywords and the weekly
date window. It does not send report drafts, internal email addresses, source
files, credentials, or unpublished company data.

Leave `approve_public_model_search` set to `false` unless the person operating
the workflow is authorized to send those public model and maker names to Bing.
When newly discovered models are appended to the CSV, the outbound list grows.
Review and renew authorization as required by your organization.

## Folder map

- `Run-WeeklyAIReport.ps1` — portable weekly entry point.
- `local_config.example.json` — copy and personalize locally.
- `requirements.txt` — Python packages.
- `data/` — watchlist, weekly input, tracker, memory, plans, and ledgers.
- `scripts/` — generator, research, reconciliation, website, and Git tools.
- `site/index.html` — maintained website file.
- `training/` — complete Codex production runbook.
- `output/` — generated reports and fallback email drafts.

## Common failures

- **Python module missing:** rerun the `pip install` command.
- **No Outlook draft:** open classic Outlook, confirm the default profile, and
  rerun. Inspect `output/weekly_ai_report_email_draft.txt` if created.
- **Research coverage mismatch:** regenerate the plan and ledger from the live
  CSV; never hand-edit counts.
- **Git authentication failure:** sign in through Git Credential Manager and
  verify write access to the repository.
- **Git non-fast-forward:** stop and reconcile the remote change deliberately.
- **GitHub Page is stale:** verify the commit is on the configured branch, then
  wait for Pages deployment and refresh with a cache-busting query.
