Test 1
- PNs in PO not found in BOM -- would be nice to have the description shown but that might mean extracting more from the PDF... the description is the characters which comes after the `|` character.
- I'm not sure the colour code is needed right now, but it prob doesn't hurt that it's there. I think the template excel MRR file has a formula built in to read the material column and then compare and return the colour code from the Data sheet tab. Should be fine to leave what you have right now.
- can we do a reset button that clears everything in the GUI to start fresh?

Test 2
- the extracted PN + Definition should have been "1027730 | IMPACT TEST COUPON" ... everything beyond that such as "IMPACT TEST COUPON - 2.0000" NOM X 24" X 6" SA516-70N. 25 Deg Bevelon 24" side." would be from another field. You also had gathered a date and dollar amounts. You'll need to look at the PDF again to make sure you only extract what you're supposed to and figure out the formatting and how to discern it.
- 

V2 feedback
- the purpose of the row splitting is user-judgement based. In the background they may receive a part number in different "heats" (for steel) and those need to be separated based on the Qty of heats on the paperwork with the material being received. So the user would need to just enter a number in a column that would essentially duplicate that row into # of rows and then they'd manually edit the Qty per row. This will ensure they have the control to override based on whatever the material split is based on how the material showed up across >1 heat. It might come in as one heat and so no extra row with the same pn is required.
- we also need a function where the user can "remove" a row from consideration for mapping/populating to the template sheets. So if I have a row that I decide based on the paperwork should not be included in the population of sheets, then it should be excluded.
- are we creating a copy in a separate sheet of all the output from V1? It would be nice to preserve that by adding a sheet at the end of the workbook as the deposit for the table that currently is the result of V1. It would include however all rows that were "split" and it would include rows that were designated as NOT to be mapped/copied/populated into the template sheets (ie it would still be there in the table but it would have been flagged as not to be mapped or whatever you call it).
- is it possible to duplicate sheets from the template file if your count runs over 48? you just add more until you meet the criteria?
- do you see any issues / have any questions or areas to push back on?

V2 feedback
- you can look at the MRR template file to find the dropdowns that are there
- 


Feedback for PDM File Procedure:
- 