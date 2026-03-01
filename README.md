![lp.png](lp.png)

# 📋 LabelPress

A open-source Google Sheets add-on that generates printable Avery label PDFs directly from your spreadsheet data. Select a template, map your columns with a flexible template syntax, and get a print-ready PDF in seconds.

## Why
I vibe coded/wrote this because I hate paywalls and limits. My neighbors don't need to pay 7$ for a plugin to make some labels for their christmas cards... Get real.

## Features

- **7 built-in Avery templates** — 5160, 5161, 5162, 5163, 5164, 5167, 8160
- **Custom label sizes** — define your own dimensions for any label sheet
- **Flexible template syntax** — use `{{Column Name}}` placeholders to design your label layout
- **Combine fields on one line** — `{{City}}, {{State}} {{Zip}}` just works
- **Smart blank line removal** — empty fields don't leave gaps on your labels
- **Skip labels** — start printing at any position for partially-used sheets
- **PDF output** — print-ready, plus an editable Google Doc if you need to tweak

## Quick Start

### Install from source

1. Open your Google Sheet
2. Go to **Extensions → Apps Script**
3. Replace the contents of `Code.gs` with [`Code.gs`](Code.gs) from this repo
4. Click **+** next to Files → **HTML** → name it `Sidebar`
5. Paste the contents of [`Sidebar.html`](Sidebar.html)
6. *(Optional)* Check "Show appsscript.json manifest" in Project Settings, then replace it with [`appsscript.json`](appsscript.json)
7. Save and reload your sheet
8. Click **📋 Label Maker → Generate Labels...**

### Usage

1. **Prepare your sheet** — headers in row 1, data starting row 2:

   | Name | Address | City | State | Zip |
   |------|---------|------|-------|-----|
   | Jane Doe | 123 Main St | Springfield | IL | 62701 |

2. **Pick a template** from the dropdown (or choose Custom)

3. **Design your label** in the template editor:
   ```
   {{Name}}
   {{Address}}
   {{City}}, {{State}} {{Zip}}
   ```

4. **Click Generate PDF** — get links to both the PDF and editable Google Doc

## Supported Templates

| Template | Labels/Sheet | Label Size | Use Case |
|----------|-------------|------------|----------|
| 5160 | 30 | 2.625" × 1" | Standard address labels |
| 5161 | 20 | 4" × 1" | Wide address labels |
| 5162 | 14 | 4" × 1.33" | Tall address labels |
| 5163 | 10 | 4" × 2" | Shipping labels |
| 5164 | 6 | 4" × 3.33" | Large shipping labels |
| 5167 | 80 | 1.75" × 0.5" | Return address labels |
| 8160 | 30 | 2.625" × 1" | Inkjet address labels |

## Custom Templates

Select **"✏️ Custom label size..."** from the dropdown to enter your own dimensions. You'll need:

- Label width & height (inches)
- Columns & rows per sheet
- Page margins (top & side)
- Gaps between labels
- Font size

These specs are usually found on the label packaging or the manufacturer's website.

## Template Syntax

Labels are defined with `{{Column Name}}` placeholders that match your sheet's header row:

```
{{Name}}
{{Company}}
{{Address}}
{{City}}, {{State}} {{Zip}}
```

- Each line in the template = a line on the label
- Multiple fields on one line are supported
- Lines that are empty after data substitution are automatically removed
- Static text (commas, dashes, etc.) is preserved

## Print Tips

- Set your printer to **Actual Size** (not Fit to Page)
- Test on plain paper first — hold it behind a label sheet to check alignment
- Use the **"Skip first N labels"** option for partially-used sheets

## Permissions

LabelPress requests minimal permissions:

| Scope | Why |
|-------|-----|
| `spreadsheets.readonly` | Read your column headers and data |
| `documents` | Create the label layout document |
| `drive.file` | Save the generated PDF to your Drive |

See [PRIVACY_POLICY.md](PRIVACY_POLICY.md) for full details. No data is stored, shared, or sent to external servers.

## Curious about ASCII Logo
[Cool Tool Here](https://patorjk.com/software/taag)

## License

WTFPL
