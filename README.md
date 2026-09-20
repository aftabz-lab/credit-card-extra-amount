# Credit Card Extra Amount

Shwapno payment-cost dashboard. Static files: no build step required.

Enable **Settings → Pages → Deploy from a branch → main → / (root)**.

Dashboard address: https://aftabz-lab.github.io/credit-card-extra-amount/

Open **Data & snapshot**, reuse the existing Zone Distribution Google and publisher connection, read both current files, review mappings, then **Take & publish snapshot**.

The initial reference is the supplied 1–29 March 2026 workbook. Its leadership is unverified until the current Zone Distribution file is loaded.

Known column aliases are detected automatically. Use **Column rules** for unfamiliar renamed headings. Use **Outlet mapping** for manual assignments or reversible exclusions.

The source target is BDT 350,000. Populated outlet targets total BDT 345,157.09; the difference is displayed under validation.

All shared writes require your existing publisher account. Automatic snapshots also run in-browser while the owner dashboard or portal is open and Drive authorization is active.

## Fully automatic snapshots (no browser tab required)

A GitHub Actions workflow (`.github/workflows/auto-sync.yml`) runs `sync-worker/auto-sync.mjs` on a schedule (every 10 minutes) so the snapshot updates whenever the source file in Drive changes, even if nobody has the dashboard open. It reuses `core.js` unmodified, so figures never drift from the manual "Read latest files" flow.

Set these two repository secrets under **Settings → Secrets and variables → Actions**:

- `GOOGLE_SERVICE_ACCOUNT_JSON` — the full JSON key of a Google service account that has been shared as a **Viewer** on the source Drive folder (`Share` → paste the service account's `...@...iam.gserviceaccount.com` email).
- `SUPABASE_SERVICE_ROLE_KEY` — from Supabase **Project Settings → API**. Used only server-side; never exposed to the browser.

The workflow can also be run on demand from the **Actions** tab (`Run workflow`). It preserves any manual outlet mapping, exclusions, and column rules already published — it only replaces the Credit Card and Zone Distribution source data when a file's signature actually changes, and refuses to publish (leaving the previous snapshot in place) if the Extra Amount / channel-sum mismatch warning is present, the same safety check the dashboard itself applies.

## Instant updates when a Drive file changes

GitHub runs scheduled workflows late (in practice every few hours, not every 10 minutes), so the schedule above is only a backup. The instant path is a small Google Apps Script (supplied separately as `drive-instant-trigger_apps-script.zip`) that runs under the Drive owner's Google account:

1. Every minute it lists the source folder and its subfolders (names, sizes and modified times only — it never opens or changes a file).
2. When any spreadsheet is added, replaced, edited, renamed or removed, it starts this workflow through `workflow_dispatch`, passing the changed file name as `reason` (shown as the run name in the Actions tab).
3. The workflow publishes the snapshot with the unchanged `core.js` rules, and every open dashboard loads it within about a second.

Typical time from a file landing in Drive to the dashboard updating: about 1–2 minutes. The script needs one script property, `GITHUB_TOKEN`: a fine-grained GitHub token limited to this repository with **Actions: Read and write**.
