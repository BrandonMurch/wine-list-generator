A Google App Script, that reads a list of selected wines from a Google Sheet, and then populates a Google Doc template. 

Simply run the createWineList() method to create.

Google sheets must have the following column headers:
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

All placeholders are wrapped in {{}}.
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
