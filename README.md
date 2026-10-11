# Family Tree website

Interactive family tree driven by `data/Master_Family_Tree_Database.xlsx`.

## Run it in VS Code
1. Open the `family-tree` folder in VS Code.
2. Install the **Live Server** extension, right-click `index.html` -> **Open with Live Server**.
   The page reads the Excel file directly, so after you edit and save the spreadsheet just refresh the browser.
   (An internet connection is needed once per load for the Excel-reading library from cdnjs.)

Double-clicking `index.html` also works; it then uses the embedded copy in `js/data.js`
(browsers block reading local files). Or click **Open Excel** in the page to load any .xlsx.

## After changing the spreadsheet (only needed for the double-click mode)
    pip install openpyxl
    python tools/build_data.py

## Excel columns used
Individual_ID, Parent_ID, Generation, Gender, Full_Name, Nick_Name, Fathers_Name, Mothers_Name,
Spouse_Name, Spouse_Father_Name, DOB, Place_of_Birth, Qualification, Date_of_Death, Place_of_Burial, Source_File.
Tree links come from Parent_ID; the generation shown is calculated from the tree position.
Sheet name should be `Family_Data` (the first sheet is used if not found).

## Using the site
- Hover a card for a quick summary; click it for the full details panel (relatives are clickable).
- Drag to pan, scroll to zoom, +/- and fit buttons, expand/collapse a branch with the button under a card.
- Filters: search, generation, gender, living/deceased, branch, place of birth, qualification, "record has".
  Filtering hides everyone who doesn't match; choosing a Branch starts the tree at that person.
- Tree/Table switch shows every column for everyone; click a column to sort.
