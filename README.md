# Investing Tracker

This repo tracks the portfolio/options workbook and the decision journal used to review trade ideas, entries, exits, rolls, and lessons learned.

## Repo Structure

- `outputs/options_tracker/Options_Trading_Tracker.xlsx` - current spreadsheet tracker.
- `.codex_work/build_options_tracker.py` - source script used to rebuild the tracker.
- `decision-journal/` - Markdown trade decision notes and post-mortems.
- `templates/` - reusable Markdown templates.

## Working Rhythm

1. Before a trade, create a decision journal entry from the template.
2. Record the thesis, position context, chart read, risk, and exit plan.
3. After the trade is entered, update the tracker.
4. Commit the tracker and journal entry together when they describe the same decision.
5. When the trade is closed or rolled, update both the tracker and the journal outcome.

## Commit Style

Use short, plain commit messages:

- `Add FIVN decision journal entry`
- `Update tracker through June 9 broker statements`
- `Close SHLS calls and record lesson`

