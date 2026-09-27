# Consultancy workspace

Open **Workspace** from the customer sidebar. Add a customer, then name a project and choose an existing parent folder on the connected environment. T3 Code creates a folder named `<project name>_YYYY-MM-DD` inside it, with `bronnen`, `word`, `excel`, `powerpoints`, and `oplevering` folders. Create a workbook with **New Excel workbook**, a Word document with **New Word document**, or a presentation with **New presentation** when you need one. Keep source material in `bronnen` and finished work in `oplevering`.

Temporary chart data and other working files belong in `.werkbestanden`. The workbench hides that folder, empty-folder markers, and editable presentation source files from the project file list.

Open a file from **Project files** to work on it. Workbooks fill the screen; choose **Ask about file** to open a side chat for row counts, missing values, duplicates, and numeric summaries of the selected sheet. PDFs, images, and text files also open on the project page. Use **Refresh** when files in the folder change. Choose **Ask AI about this project** or **New chat** for broader questions about the project files.

Word documents open in a page editor with Save, Download, and a document chat. Each document keeps its chat when you reopen it. The editor supports text, headings, bold, and italic. Documents with tables, images, or formatting that cannot be preserved open as read-only text previews; download the original to see its full layout. Use **New Word document** to create a file with the company theme and section order. An admin can change both under **Settings > Company**. Use **Apply company standard** in an open document to apply the current theme and add missing sections.

Workbook editing starts off. Turn on **Edit**, then open **Ask about file** to allow the whole workbook or choose a worksheet, cells, rows, or columns. Click any cell to select its full row or column, then use the controls there to add or remove rows and columns or create a worksheet. Double-click a cell within the chosen scope to change it, then select **Save**. Workbooks with formulas or Excel features that cannot survive structural edits keep the row, column, and worksheet controls disabled. Questions from the file chat reuse that workbook's agent thread and tell the agent which cells it may change without showing the workbook context in your message. Turn **Edit** off to send read-only questions.

Files added through the earlier upload flow remain under **Previously added files**. Their originals are stored in the browser where they were added.

**Project AI usage** shows input and output tokens by model for AI turns completed in that project after tracking was enabled. Use **Refresh** after a chat ends. Cost is an API price estimate; subscription charges can differ, and models without a known price appear as unpriced.
