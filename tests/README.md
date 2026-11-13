# Tests

Tests for the Online Helpdesk Apps Script project. They run on a PC with Node.js — nothing here is deployed to Apps Script.

## What is covered

| Folder | What it checks | How |
| --- | --- | --- |
| `server/` | `.gs` code: Order Project rules, the helper's data, form pages for more than 8 items, shared helper sets | The `.gs` files run in a Node sandbox with a small fake of the Google services |
| `browser/` | `index.html`: the Order Project helper (Card holder, Strip marker, Cam switch plate, Annunciator, Ammeter / Voltmeter, fonts), PDF merge, Gantt export, Admin-only file paths | The page runs in Chromium with a fake `google.script.run` (`mock/head.html` + extra mock data) |

The browser scripts print what they see and fail only when the page throws an error. Read `tests/out/*.log` and the screenshots in `tests/out/` to check the result by eye.

## Run

```bash
cd tests
npm install            # once: pdf-lib + playwright
npx playwright install chromium   # once, if Chromium is not installed yet
bash run.sh
```

`run.sh` checks the `.gs` syntax, runs the server tests, builds the test pages into `tests/out/`, then runs the browser tests. It ends with `ALL PASSED` or `SOME FAILED`.

## Notes

- `mock/head.html` holds made-up data only (no real file list from the file server).
- These tests cannot reach Google: exporting a real order (PDF / Excel from the template) still has to be tried once on the deployed web app after a change to `OrderService.gs`.
- Add a test next to the feature it covers: `server/<name>.test.js` must print `N passed, 0 failed` on its last line; `browser/<name>.js` must exit 0.
