# Broga TEDUH daily collection package

## Files
- `index.html`: dashboard v28 with a shared-history loader.
- `scripts/collect_teduh.mjs`: validates the configured unit scope and collects dated TEDUH snapshots.
- `.github/workflows/daily-teduh-collection.yml`: daily GitHub Actions schedule at 08:30 Malaysia time (00:30 UTC), plus manual runs.
- `data/teduh-history.json`: shared history file; the workflow appends/updates one record per development per Malaysia date.

## Install into the existing repository
Upload/replace these paths at the root of `broga-competitor-sales-dashboard`:
1. Replace `index.html` with the included file.
2. Add `.github/workflows/daily-teduh-collection.yml`.
3. Add `scripts/collect_teduh.mjs`.
4. Add `data/teduh-history.json`.

Keep the repository Pages source set to `main` and `/(root)`. The dashboard loads shared records from `./data/teduh-history.json`.

## First test
After committing these files, open Actions → Daily TEDUH collection → Run workflow. The job saves data only if all 12 configured endpoints return valid JSON and all strict scope checks pass. This intentionally prevents partial or mismatched snapshots from being treated as valid history. If a scope check fails, read the failing endpoint and error in the job log before changing any rule.

## Scope guardrails
- Springvale: L1-L27 and L28-L335 only (335 total); L336-L373 excluded unless separately reviewed and explicitly configured.
- Birchpark: terrace groups 57513/57514/57515/57516 only (46+56+25+25 = 152); excludes rumah bandar.
- Sapphire: rumah teres only (136); excludes rumah bandar and rumah sesebuah.
- M Legasi Phase 1: three confirmed groups 111 + 138 + 81 = 330.
- M Legasi Phase 2: group sizes 26 + 126 + 104 = 256.
- M Legasi Phase 3: group sizes 44 + 85 + 102 = 231.
- Other mapped developments use the configured comparison totals in the collector script.

## Important limitations
- No email alerts are configured in this first phase, per the chosen staged rollout. Failed runs are visible in GitHub Actions.
- Daily observations are saved to GitHub, not the local C: drive.
- The workflow does not automatically change confirmed comparison denominators when TEDUH totals change. Scope mismatches fail the run so they can be reviewed.
- Legacy observations remain embedded in the HTML. Shared daily observations are loaded from the JSON file and merged by project/date.
- GitHub scheduled workflows can be delayed. The daily schedule is a best-effort collection time, not a guaranteed exact run time.
