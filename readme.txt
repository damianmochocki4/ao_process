1. "Masterlist" tab has to be the first tab, "Los" tab - the second. Worksheets' names are not important.
2. "Masterlist" tab has to have columns with "AO_ID", "NE3_Baustatus" and "Fördertyp" headers.
3. "Los" tab has to have "AO_ID" header.
4. Output will be created in the same directory as the Excel file, with the "edited_current_date" appendix.
5. Tables have to be deactivated before running the script. If not, Excel will want to fix the output file after opening it.
6. Rows with values in "Adresstyp aus BST" column other than "Förderadresse (weißer Fleck)" will be exported to different tabs (named "Addresstyp_value") and marked with colors based on a value:
- empty - orange,
- "Vortriebsadresse" - yellow,
- any other value - light yellow.
7. Rows with values in "Baustatus NE3 aus BST" column other than "In Vorbereitung" will be exported to different tabs (named "Baustatus_NE3_value")