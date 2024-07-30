1. "Masterlist" tab has to be the first tab, "Los" tab - the second. Worksheets' names are not important.
2. "Masterlist" tab has to have columns with "AO_ID", "NE3_Bausta" and "Fördertyp" headers.
3. "Los" tab has to have "AO_ID" header.
4. Output will be created in the same directory as the Excel file, with the "_edited" appendix.
5. Tables have to be deactivated before running the script. If not, Excel will want to fix the output file after opening it.
6. Rows with values in "Adresstyp aus BST" column other than "Förderadresse (weißer Fleck)" will be exported to different tabs (named "value_date") and marked with colors based on a value:
- empty (NULL) - orange,
- "Vortriebsadresse" - yellow,
- any other value - light yellow.
