# Coralogix log exports

## Armor public-API calls -> .xlsx

One row per gateway call for a customer on one UTC day. Query lives in
`queries/armor_calls_export.dp` (`{{BUSINESS_NUMBER}}` / `{{KEY_PREFIX}}` placeholders), so the
same text can be pasted into Coralogix Explore (Archive tier) after substituting the two values.

```
export CORALOGIX_API_KEY=...            # or CORALOGIX_KEY_FILE=/path/to/key.json (apiKey.keyValue)
python coralogix/armor_export.py \
  --business-number <15-digit number> --key-prefix <first 8 chars of API key> --day YYYY-MM-DD \
  [--out out/file.xlsx] [--expected-rows N]
```

- Day is a full **UTC** day (e.g. 2026-10-11 = 11/10 05:30 -> 12/10 05:30 IST).
- Hitting 10,000 rows or a `bytesScannedLimitWarning` splits the day into 6-hour windows,
  merges, and dedupes on `armor_request_id`. A window that still hits a limit is reported as partial.
- Output: one sheet, columns `timestamp, applicationName, api, armor_request_id, client_ip,
  full_uri, method, status, time_ms, trace_id`, sorted by timestamp, header frozen, autofilter on.
- The script prints the total, an api x status count table, and any warning/error lines. It never
  prints the key. Needs `requests` and `openpyxl`.

Tests (mocked API, no network): `python -m unittest discover -s coralogix/tests`

## Armor: who is calling an account (source IPs)

For "where is this API traffic coming from" questions (e.g. V2BIRD-37848). Replace
`{{BUSINESS_NUMBER}}` and paste into Coralogix Explore (Archive tier, UTC range, <= 1 day per run
to stay under the scan limit):

- `queries/armor_account_sources.dp` - summary by client IP x key prefix x endpoint x status, with first/last seen.
- `queries/armor_account_calls.dp` - one row per call (UTC + IST time, IP chain, key prefix, trace id). Any `api_key=` query value is masked; only 8-char key prefixes are ever extracted.

Calls for child locations don't carry the parent's business number: once the summary shows the key
prefix, re-run with `|| $d.message:string.contains('<KEY_PREFIX>')` added to the business-number filter.
