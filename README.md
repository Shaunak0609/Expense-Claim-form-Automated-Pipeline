# ECF Pipeline

A Google Apps Script that turns reimbursement requests from a Google Form into filled-in
**UBC Engineering Design Teams Expense Claim Forms (ECFs)**. Built for the **2026-27** ECF template.

Making an ECF by hand used to take **15–30 minutes per request**. With this script it takes a few minutes,
most of which is checking the result.

## Why it edits the Excel file directly

The ECF is an Excel file with old-style checkboxes and dropdowns. Opening it in Google Sheets
(or saving it with most spreadsheet libraries) silently deletes them. So the script never converts it.
It unzips the original `.xlsx`, edits the XML inside, and zips it back up.
Everything in the template (checkboxes, dropdowns, formatting, formulas, logo) stays as it is.

## What it does

For each form response (one response = one claimant = one ECF):

1. **Reads** the claimant's name, email, student number and signature, plus up to 10 purchases.
   Purchases are read until one is empty or "Do you have more purchases to claim?" isn't "Yes".
2. **Checks** the response. If anything is wrong, no ECF is made: the status becomes `Needs Info`
   and the **ECF Log** tab says exactly what to fix.
   - Receipt Amount and Exchange Rate must contain exactly one number above 0.
     Text around it is fine: `100 CAD`, `CAD 100`, `$1,234.50` and `USD 45.99` all work;
     `1 USD = 1.38 CAD` is flagged because it isn't clear which number is meant.
   - Student Number must be a number; Purchase Date must be a date.
   - Expense Category must be one of the 5 ECF categories (Prototype Project, Competition - People Costs,
     Competition - Physical Logistics, Professional Development, Operational/Administrative Expenses).
     Form wordings that differ (e.g. "Prototyping Project") can be mapped in `CATEGORY_ALIASES`.
   - Full Name and Vendor can't be empty.
   - The Signature link must open a PNG or JPG image.
3. **Fills a copy of the ECF:**
   - Header: legal name, date, student #, email, design team, worktag
   - Purchase rows 1–10: date, vendor, amount, exchange rate, category, project/event, purpose
     (the ECF has 15 rows but the form collects at most 10; the Project ID column is left blank)
   - Ticks the **Missing Receipt** checkbox on purchases marked "Yes"
   - Claimant box: print name, signature image from the form upload, and date
   - Team's designated approver box: print name, signature image, and date (set once in the config)
   - Faculty supervisor box is left empty for the supervisor to sign
   - CAD amounts and the TOTAL are left to the template's own formulas
4. **Saves** `ECF - <Name> - <yyyy-mm-dd HHmm>.xlsx` to a Drive folder, sets the status to
   `ECF Generated`, and logs a link to the file.

See [sample-output-10-purchases.xlsx](sample-output-10-purchases.xlsx) for an example made from fake data.

## Files

| File | What it is |
|---|---|
| [Code.gs](Code.gs) | The script |
| [Config.example.gs](Config.example.gs) | Settings template. Copy it to `Config.gs` and fill in your values |
| `Config.gs` | Your real settings (Drive IDs, names). **Not committed**, see `.gitignore` |
| [sample-output-10-purchases.xlsx](sample-output-10-purchases.xlsx) | Example output from fake data |

## Setup (one time, about 10 minutes)

1. **Upload the blank ECF to Google Drive as an `.xlsx`.**
   First turn off Drive → Settings → "Convert uploads", or the checkboxes are lost.
   Open the file and copy its ID from the address bar: `.../d/<FILE ID>/...`
2. **Create an output folder** in Drive and copy its ID: `.../folders/<FOLDER ID>`
3. **Optional: upload the approver's signature** (PNG or JPG, cropped, on white) and copy its file ID.
4. **Add two options to the status dropdown** (column A of the responses tab): `ECF Generated` and `Needs Info`.
5. In the responses Google Sheet, open **Extensions → Apps Script** and add two files:
   - `Code.gs`: paste in [Code.gs](Code.gs)
   - `Config.gs` (click **+ → Script**, name it `Config`): paste in
     [Config.example.gs](Config.example.gs) and fill in:

     | Setting | Value |
     |---|---|
     | `TEMPLATE_FILE_ID` | ID from step 1 |
     | `OUTPUT_FOLDER_ID` | ID from step 2 |
     | `TEAM_NAME` | Your design team's name |
     | `WORKTAG` | Your team's worktag, copied exactly from the ECF list (e.g. `PM016212 EDT-Smart City`) |
     | `APPROVER_SIGNATURE_FILE_ID` | ID from step 3, or `''` to leave it out |
     | `APPROVER_NAME` | Team's designated approver, or `''` |

6. Click **Save** (no need to Deploy), then reload the Google Sheet. An **ECF** menu appears.
   The first run asks for permission to use Sheets and Drive. If you see "Google hasn't verified this app",
   choose **Advanced → Go to (project)**. That's normal for your own scripts.

## Usage

- **ECF → Generate ECF for selected row(s)**: select one or more response rows first.
  This ignores the status, so you can regenerate a row after fixing it.
- **ECF → Generate ECFs for all "Request Made" rows**: processes every pending row.

A popup summarizes each row, and the **ECF Log** tab keeps a link to every ECF or the reason a row was rejected.

**Optional schedule:** in Apps Script, **Triggers** (clock icon) → **Add Trigger** →
function `generateForPendingScheduled`, *Time-driven*, *Week timer*, *Every Saturday*.

## Things to know

- **Open ECFs in Excel, not Google Sheets.** Google Sheets can't display the checkboxes, and saving
  the file there deletes them. Download the file and open it in Excel.
- **Purchases are read by column position**, because the form's headers for purchases 2–10 aren't consistent.
  Each purchase is 13 columns wide. If questions are added or reordered *inside* the purchase section,
  update `BLOCK` in `Code.gs`. Questions added after it are fine.
- **Run the script from the form owner's Google account.** Form uploads (signatures) belong to that account.
- **Signatures go in exactly as uploaded.** A grey background or a sideways photo will look that way on the ECF.
  iPhone HEIC photos aren't supported.
- **Check the signing rules.** The 2026-27 ECF says pasted signature images are not accepted
  (sign the PDF electronically or approve by email instead). Confirm with your finance contact before
  relying on the signature images; set `APPROVER_SIGNATURE_FILE_ID` to `''` to leave the approver box for a real signature.
- The script writes `.xlsx` only. Export to PDF from Excel if you need one.
- **New ECF version?** Cell positions live in `ECF` at the top of `Code.gs` and categories in `Config.gs`.
  If the finance office changes the layout, those are the places to update.

## How it works (for developers)

- `run_()` reads the sheet, then for each row calls `parseClaim_()` (pure parsing and validation)
  and `loadSignature_()` (Drive), then `buildEcf_()`.
- `buildEcf_()` unzips the template with `Utilities.unzip`, passes the XML text to `fillEcfXml_()`,
  and re-zips the changed parts.
- `fillEcfXml_()` is pure string-in, string-out, so it can be tested in Node without Google:
  - `setCell_()` replaces existing template cells, keeping their styles
  - `tickCheckboxes_()` finds each checkbox's row from its anchor, then sets `checked` in its
    `ctrlProp` file and `<x:Checked>` in the VML drawing
  - `addSignatureImage_()` adds a picture anchor to the drawing, a relationship and a media file
  - it sets `fullCalcOnLoad` so Excel recalculates the formulas on open, and writes cached values
    so previews show the right totals
