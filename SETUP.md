# Credit Card Extra Amount: Outlet Detail update

## Upload these four files

1. Extract `Credit_Card_Outlet_Detail_Fix_GitHub_Upload.zip`.
2. Open your existing GitHub repository **aftabz-lab/credit-card-extra-amount**.
3. Upload these files at the repository's top level:
   - `app.js` — replace the existing file.
   - `index.html` — replace the existing file.
   - `outlet-detail-layout.js` — new file.
   - `outlet-detail-layout.css` — new file.
4. Commit the upload to the branch used by GitHub Pages and wait for deployment to finish.
5. Open https://aftabz-lab.github.io/credit-card-extra-amount/ and press Ctrl+Shift+R once.

You do not need to upload this instruction file. No other repository files need replacement.

## Requested changes only

- Remove the second, placeholder **Outlet name / Unnamed outlet** column from **Outlet Detail**.
- Remove **Target** and **Save amount** from the same display and its Columns menu, including when an older browser preference selected them.
- Retain the actual **Outlet Name** column from the raw data and all other existing columns.
- Freeze the **Outlet Detail** heading when it reaches the top of the browser window. Keep the table column headers directly below the frozen heading.
- While the heading is frozen and the table is wider than the available space, show a synchronized horizontal scrollbar at the bottom of the browser window. Scrolling either this bar or the table moves the same columns. The bar hides outside Outlet Detail and when horizontal scrolling is unnecessary.

Existing calculations, data, filters, sorting, pagination, downloads, fonts, themes, other dashboard sections, Google Drive logic and snapshot rules are preserved. Target and saving remain in the backend and existing calculation/export logic; this change removes only their Outlet Detail display.

## Verification

Checked using the existing published snapshot with **1,020 eligible outlet records**. Before/after backend records, financial facts and snapshot time matched. Confirmed only the three requested table columns disappeared, actual outlet names remained, frozen heading and column headers aligned, two-way horizontal scrolling worked, the scrollbar resized when columns changed, and the layout worked in a narrow window and both themes. No browser script errors occurred.
