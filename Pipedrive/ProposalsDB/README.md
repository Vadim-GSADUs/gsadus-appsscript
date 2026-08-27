# ProposalsDB — SUPERSEDED (2026-08-27)

> **This Apps Script project is retired.** PP# minting, Drive proposal-folder
> creation, and the CRM PP#/Folder-URL fields are owned by the **WebApp
> proposal engine** (`C:\GSADUs\WebApp\lib\proposals\`, live in prod since
> 2026-08-27; spec + amendments in
> `Documents\GSADUs-planning\2026-08-19-proposal-engine\`). The ledger of
> record is the `proposals` schema in the `gsadus-web-catalog` Supabase
> project.
>
> Decommission record (spec §12.2, executed 2026-08-27, §12.1 gate waived by
> owner at go-live day):
>
> - Triggers deleted + webhook deployments archived 2026-08-25; re-verified
>   2026-08-27 (spreadsheet untouched since the freeze).
> - The `PP{n} [PLACEHOLDER]` folders are gone — consumed as sanctioned
>   instant-mint fallbacks (Amendment A3); their numbers stay absorbed in the
>   engine's number registry.
> - The gsheet's Logs tab (1,451 rows) is exported to
>   `Documents\GSADUs-planning\2026-08-19-proposal-engine\legacy-proposalsdb-logs.csv`.
> - The spreadsheet is renamed
>   `ProposalsDB [ARCHIVED 2026-08-27 - superseded by WebApp proposal engine]`
>   and frozen in place.
>
> Do not re-enable triggers, re-deploy the webhook, or extend this code. The
> sources remain for historical reference only (clasp-managed archive).
