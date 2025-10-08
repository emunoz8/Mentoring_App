# Code Review Suggestions

## QueueScript.html
- **Avoid double-encoding `ids` query parameter.** `buildIndividualNotesUrl` currently joins the IDs, manually wraps the result in `encodeURIComponent`, *and* pushes it through `URLSearchParams`. This encodes commas as `%252C`, breaking the downstream parser. Drop the manual `encodeURIComponent` call and let `URLSearchParams` handle the encoding. Relevant snippet:
  ```js
  if (ids?.length) params.set('ids', encodeURIComponent(ids.join(',')));
  ```
- **Remove the duplicated function declaration fragment.** `renderBoard` still contains a stray comment fragment (`// rows expected: [{ id, nfunction renderBoard(rows){`) that duplicates the opening of the function signature and risks confusing future edits. Clean up the comment so the function body starts cleanly.
- **Address unused helpers.** `statusClass` and `statusTitle` are defined but never referenced. Consider wiring them into the UI (e.g., for status styling/tooltips) or removing them to keep the client bundle trim.

## QueueServer.js
- **Batch Apps Script writes.** `claimRows` and `markProcessed` loop over each row, invoking `getRange(...).setValue(...)` multiple times per iteration. This generates 3–4 client/server round-trips per row. Replace the per-cell writes with a single `getRangeList().setValues(...)`/`getRange().setValues(...)` batch write (possibly after building an array of updates) to reduce execution time and quota usage.
