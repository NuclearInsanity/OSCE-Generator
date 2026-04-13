# AFRM OSCE Picker

Browser-based app for drawing random OSCE stations from an Excel workbook of past stations.

## Local use

From the project folder, run:

```bash
python3 osce_picker_app.py
```

Or launch:

```bash
./Launch\ AFRM\ OSCE\ Picker.command
```

That regenerates:

- `AFRM OSCE Picker.html` for local browser use
- `docs/index.html` for GitHub Pages

## What it does

- Reads the workbook directly from `.xlsx` without `pandas` or `openpyxl`
- Switches between the official OSCE station bank and the mock/practice station bank
- Filters by year, type, main topic, sub-topic, and specific examination
- Keeps a carousel over the currently filtered station set so you can move left and right through matches
- Keeps random selection from the filtered set and from the full dataset
- Shows the original source question as the main station prompt
- Reveals the polished question prompt when you click `Show question`
- Reveals the GPT answer when you click `Show answer`

## GitHub Pages

This project can be published using GitHub Pages from the `docs/` folder.

1. Push the project to a GitHub repository.
2. In the repository, open `Settings` > `Pages`.
3. Set the source to `Deploy from a branch`.
4. Select the main branch and the `/docs` folder.
5. Save and wait for the site URL to be generated.

When the workbook changes:

1. Run `python3 osce_picker_app.py --no-open`
2. Commit the updated `docs/index.html`
3. Push the changes to GitHub

## Notes

- If the workbook moves or is renamed, update the `DEFAULT_WORKBOOK` path in `osce_picker_app.py` or run the script with `--workbook "/path/to/file.xlsx"`.
- If the practice workbook moves or is renamed, update the `DEFAULT_PRACTICE_WORKBOOK` path in `osce_picker_app.py` or run the script with `--practice-workbook "/path/to/file.xlsx"`.
- If a station has no polished question, the app falls back to the original question text in the question panel.
- Anything embedded in `docs/index.html` is published content once the GitHub Pages site is live, including the mock/practice station bank.
