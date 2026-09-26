---
name: excel-chat-visuals
description: Turn Excel workbook data into useful charts and tables shown in chat. Use when someone asks to understand, compare, or summarize data from an Excel file, especially trends, costs and income, or department and project counts.
---

# Excel chat visuals

Use this skill to answer a question about workbook data with a visual in the conversation. The goal is a chart or table the reader can understand at a glance, with a brief supported takeaway.

## Workflow

1. Inspect the workbook before choosing a visual. Identify the relevant sheet, headers, date or category fields, measures, units, currency, and any filters or duplicate records that affect totals. Use the workbook-reading tools available in the current environment. For generic spreadsheet-file analysis, follow the available Spreadsheets skill; for an explicitly selected live Excel session, follow Excel Live Control.
2. Map the user's question to fields. If a necessary mapping is ambiguous, ask one focused question rather than guessing. Keep source values and calculated values distinct.
3. Aggregate only as needed. Group time-series values into sensible periods, and count distinct projects when the question asks for projects rather than rows. Preserve missing values as missing; do not turn them into zero. State material assumptions such as date range, grouping, or distinct-count logic.
4. Choose the simplest visual that answers the question. Use a line chart for change over time, a grouped bar chart for comparable measures such as cost and income by period, a sorted horizontal bar chart for department counts, and a Markdown table when exact values or many categories matter more than shape. Use the fill-in templates in `assets/visual-templates/` as starting structures, not as a required style.
5. Show the visual in chat using the visualization output method available in the current environment. If the current environment has the `visualize` skill, follow its full instructions for inline visuals and include its required content reference in the final response. If no inline visualization surface is available, show a readable Markdown table or concise text chart directly in chat.
6. Add a short takeaway grounded in the displayed data. Include units, period, and source sheet or field names where helpful. Flag incomplete or ambiguous data that could change the result.

## Visual rules

- Compare like with like: align time periods, currencies, and units before plotting. Do not combine unlike currencies or units on one scale.
- For cost versus income, plot both series on the same period and scale when they share a unit. If useful, add net income as a clearly labeled derived series, calculated as income minus cost.
- For category counts, sort by count and show the full set when it remains readable. If there are many categories, show the leading categories and state how many are omitted, or use a table.
- Label axes, series, units, and date range. Keep number precision appropriate to the source and question.
- Do not infer causation, fill missing values, or imply forecasts from historical data. Distinguish actuals, estimates, and forecasts when the workbook does.
- Do not edit the source workbook unless the user asks for an edit.

## Fill-in templates

Use the JSON examples under `assets/visual-templates/` to organize chart data and labels before rendering. Replace every `<...>` value with inspected workbook information. Remove fields that do not apply, and add only fields needed to make the visual clear. These files describe the data and presentation; render them with the current chat's available visualization surface rather than returning raw JSON as the answer.
