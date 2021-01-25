A Google App Script, that reads a list of selected wines from a Google Sheet, and then populates a Google Doc template.

Before running, a folder must be created:

1. Right click in Google Drive.
2. Click "New folder", and type in an appropriate name.
3. Enter into the folder and get the id in the address bar, it will be the string after "folders/"
4. Enter this id into the top of the script, on the line "const FOLDER_ID = ..."

Next Create a template:

1. Create the template following the necessary rules below, or copy WineListTemplate.docx in the drive.
2. Get the Id of this document in the address bar, it will be the after "/document/d/"
3. Enter this id into the top of the script, on the line "const TEMPLATE_ID = ..."

To run, add a menu item to the spreadsheet:

1. In the spreadsheet, click Tools > Script editor.
2. Copy code from addMenuItem.js to the script editor.
3. Run onOpen().

Google input sheet must have the following column headers:

- Name
  - vintage (with 4 digit year)
  - name within single quotation marks)
  - Any other information between parenthesis ()
- Grapes
- Restaurant Price
- Type
- Country
- Region
- Producer
- Hide From Wine List
  - "True" will hide the wine from the created wine list

All template placeholders are wrapped in {{ }}.
Current accepted placeholders are:

- category
- category_maceration
- region
- cuvee
- grapes
- cuvee_maceration
- price

Country in the sheet will be replaced by an image of rotated text on the left hand side.
The current country options are:

- Australia
- Austria
- Germany
- Hungary
- Italy
- Japan
- South Africa
- Spain
- USA
