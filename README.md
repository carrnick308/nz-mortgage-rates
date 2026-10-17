# NZ mortgage rates – automatic monthly update

Checks the Reserve Bank's mortgage rate data every Monday morning: B20
standard rates and B21 special rates. When RBNZ has released a new month,
it rebuilds:

- `docs/historical-mortgage-rates.xlsx` – the Historical mortgage rates spreadsheet
- `docs/nz-mortgage-rates.json` – the data and headline figures, for the app
- `docs/nz-mortgage-rates.csv` and `docs/nz-mortgage-rates-b21.csv` – plain data
- `docs/index.html` – a simple web page with the table and download links

Once GitHub Pages is on, these live at permanent links:

    https://carrnick308.github.io/nz-mortgage-rates/historical-mortgage-rates.xlsx
    https://carrnick308.github.io/nz-mortgage-rates/nz-mortgage-rates.json
    https://carrnick308.github.io/nz-mortgage-rates/

## Before you start

- **Make the repository public.** GitHub Pages is free only for public
  repositories. Nothing private is in here; it's RBNZ's published data.
- **Use exactly the name `nz-mortgage-rates`**, or the links above change.
- **Keep the `.github` folder** when copying files in. Windows can treat
  folders starting with a dot as hidden, and without it nothing runs.

## One-time setup (about 15 minutes)

1. On github.com, create a new **public** repository called
   `nz-mortgage-rates`. Don't add a README.
2. In GitHub Desktop, clone it to `C:\Projects\nz-mortgage-rates`.
3. Unzip this folder's contents into it, including `.github`, `assets`,
   `docs`, `inbox` and `scripts`. Commit ("Initial setup") and push.
4. On github.com, open the repository → **Settings → Actions → General** →
   *Workflow permissions* → choose **Read and write permissions** → Save.
5. **Settings → Pages** → *Source: Deploy from a branch* → Branch **main**,
   folder **/docs** → Save.
6. **Actions** tab → **Update mortgage rates** → **Run workflow**. (Pushing
   in step 3 may already have started a run. If that run failed before
   step 4, ignore it and use this one.)

## Reading the result of the first run

- **Green tick:** automation works. Open the spreadsheet link above and check
  it. Then, on your Squarespace page, change the "Historical mortgage
  rates" link to the `.xlsx` link above. You're done: it now updates itself.
- **Red cross saying RBNZ returned a web page or couldn't download:** RBNZ's
  bot protection is blocking GitHub. Use the fallback below. It turns the
  monthly job into about two minutes.
- **Red cross saying columns couldn't be found:** RBNZ's file layout differs
  from what the script expects. Send the error message to Claude.

GitHub emails you whenever a run fails, so a broken month won't go unnoticed.

## Fallback: monthly two-minute upload

1. Download `hb20.xlsx` from the
   [RBNZ B20 page](https://www.rbnz.govt.nz/statistics/series/exchange-and-interest-rates/new-residential-mortgage-standard-interest-rates)
   and `hb21.xlsx` from the
   [RBNZ B21 page](https://www.rbnz.govt.nz/statistics/series/exchange-and-interest-rates/new-residential-mortgage-special-interest-rates).
2. On github.com, open the `inbox` folder → **Add file → Upload files** →
   drop in both files together → Commit.

That upload starts the update straight away. The files are removed from
`inbox` after they're used.

## What's in the spreadsheet

- **Summary** – the four tables visitors look at: standard rates since
  January 2005 (average, highest, lowest, range); the average over periods
  of 10 years or longer; how often each term has been the cheapest; and
  special rates since January 2017.
- **Standard rates / Special rates** – month by month, with each month's
  cheapest rate highlighted and a "Cheapest term" column.
- **Workings** – one row per period, starting every January from 2005.
  A new start year is added automatically once it has 10 years of
  data: January 2017 joins when December 2026 data arrives.

Every figure is a live formula, so visitors can click a cell and see how
it's calculated. Ties go to the shorter term.

## Safety checks

The published files are left untouched if the new data:

- is missing a term (floating only applies to standard rates)
- has rates outside 0–30%
- has fewer months than what's already published (stops a broken RBNZ file
  wiping your history)

## Using it in the app

Fetch the JSON link above, cache the last copy on the device, and show
"Rates to <latest_month>". `summary` holds the headline figures, so the app
doesn't need to calculate them:

```json
{
  "summary": {
    "latest_month": "2026-09",
    "standard": {"1y": {"average": 6.004, "high": 9.9, "low": 3.17, "range": 6.73}},
    "standard_10yr_plus_average": {"1y": 5.421},
    "cheapest_term_share": {"1y": 61.09, "6m": 20.29}
  },
  "series": {
    "B20": {
      "latest_month": "2026-09",
      "terms": [{"key": "floating", "name": "Floating"}, {"key": "1y", "name": "1 year"}],
      "months": [{"month": "2026-09", "floating": 5.43, "1y": 4.89}]
    }
  }
}
```
