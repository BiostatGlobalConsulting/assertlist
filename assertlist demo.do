* Examples of using assertlist, assertlist_cleanup and assertlist_replace on Stata's famous auto dataset
* Dale Rhoda
* March 20, 2020

* Make sure you are cd in the location you want to run your test as does create output.
********************************************************************************
* Multiple excel and .do files are created during this demo.
* We want to first wipe out any that may already be existing in the current directory
* For this demo we want to erase any old excel files 
* Start with an empty Excel files
foreach f in al_xl_demo al_xl_demo_not_cleaned al_xl_demo_clean al_xl_demo_clean_and_sorted al_xl_demo_org {
	capture erase `f'.xlsx
}

* We also want to remove any old do files that are created later in this program
foreach f in replacement_commands al_xl_demo_replace_commands al_xl_demo_replace_commands_non_cleaned al_xl_demo_clean_replace_commands al_xl_demo_clean_replace_commands_with_comment {
	capture erase `f'.do
}

* Open Stata's auto dataset
sysuse auto, clear
********************************************************************************
* First we will take a look at rep78. This will show there are 5 missing values.
tab rep78, m

* Run assertlist to show the line number for each contradiction to the screen
assertlist !missing(rep78)

* Instead of the line number we want to see make and rep78 for all contradictions
assertlist !missing(rep78), list(make rep78)

* You can include as many variables as you would like in the list option.
assertlist !missing(rep78), list(make price mpg rep78 headroom trunk foreign)

********************************************************************************
* If you would like to save the list of contradictions, use the excel option

* With no list option, the line numbers of rows that fail assertion are exported
assertlist !missing(rep78), excel(al_xl_demo.xlsx) sheet(do_not_want_to_correct)

* In this case we list the make and the value of rep78, and we include an informative tag
* This will add the new variables to the far right of the tab
* And the tag will be included in both the Assertlist_Summary and do_not_want_to_correct tabs.
assertlist !missing(rep78), excel(al_xl_demo.xlsx) sheet(do_not_want_to_correct) list(make rep78) tag(Missing value for rep78)

* If we want to be able to go back and put in a corrected value specify the FIX
* This creates extra spreadsheet columns for all variables provided in CHECKLIST
* The sheet will include "_fix" at the end

* If the user populates these columns with the correct value they can use the ASSERTLIST_REPLACE program
* to read in these values and make the replacements.
assertlist !missing(rep78), excel(al_xl_demo.xlsx) sheet(want_to_correct) fix idlist(make) checklist(rep78) tag(Missing value for rep78)

* Lets rerun the previous test but add a second variable to IDLIST.
* This will ERROR out because the IDLIST provided is a different IDLIST then used in the previous line but has the same sheet name.
assertlist !missing(rep78), excel(al_xl_demo.xlsx) sheet(want_to_correct) fix idlist(make price) checklist(rep78) tag(Missing value for rep78)

* So we can rerun this by either changing the SHEETNAME or IDLIST
* In this case we will change the SHEETNAME
* The sheet will include "_fix" at the end
assertlist !missing(rep78), excel(al_xl_demo.xlsx) sheet(want_to_correct2) fix idlist(make price) checklist(rep78) tag(Missing value for rep78)

********************************************************************************
* Assertions where all lines pass

* If excel is not specified... nothing happens as all lines passed the assertion
assertlist gear_ratio < 4, list(make gear_ratio)

* When Excel is specified
* Note that the program makes an entry in the Assertlist_Summary sheet, but does 
* not make a new sheet named 'no_fails' or no_fails_fix
assertlist inlist(foreign,1,0), excel(al_xl_demo.xlsx) sheet(no_fails) list(make foreign)
assertlist gear_ratio < 4, excel(al_xl_demo.xlsx) sheet(no_fails) fix idlist(make) checklist(gear_ratio) tag(Gear ratio higher than expected)
********************************************************************************
* Let's look at an assertion involving an 'if' statement 
assertlist headroom <= 3 if length <= 180, list(make headroom length)

* These two observations violate our expectations...so we might go and look up 
* the numbers again and possible fix the data. But we don't yet know whether we 
* will replace the value of headroom or of length, so we list BOTH in the 
* CHECKLIST option

* Note this will append to the existing want_to_correct_fix tab as the idlist is the same
assertlist headroom <= 3 if length <= 180, excel(al_xl_demo.xlsx) sheet(want_to_correct) fix idlist(make) checklist(headroom length) tag(Headroom seems large for small length)


* Finally, here is an example of sending output from several assertions to a 
* single worksheet which can be used for data cleaning.

* I don't know much about cars...probably the values in the dataset are fine.
* I have inserted some assertions here that look for large values for some
* of these variables and writes out which cars had the large values.  We will
* pretend that a data manager might go back and check the large values to 
* see whether they are correct or mis-typed.

assertlist price < 15900, excel(al_xl_demo.xlsx) sheet(several_tests) fix idlist(make) checklist(price) tag(Price seems very high)
assertlist mpg < 40, excel(al_xl_demo.xlsx) sheet(several_tests) fix idlist(make) checklist(mpg) tag(MPG seems very high)
assertlist rep78 < 5 if !missing(rep78), excel(al_xl_demo.xlsx) sheet(several_tests) fix idlist(make) checklist(rep78) tag(rep78 seems very high)
assertlist headroom < 5, excel(al_xl_demo.xlsx) sheet(several_tests) fix idlist(make) checklist(headroom) tag(Head room seems very high)
assertlist trunk < 22, excel(al_xl_demo.xlsx) sheet(several_tests) fix idlist(make) checklist(trunk) tag(Trunk space seems very high)

* Add a test with a second check variable... This will create var_2* variables
assertlist trunk < 22, excel(al_xl_demo.xlsx) sheet(several_tests) fix idlist(make) checklist(trunk mpg) tag(Trunk space seems very high)

* Add check for missing rep78 so we can add some conflicts
assertlist !missing(rep78), excel(al_xl_demo.xlsx) sheet(several_tests) fix idlist(make) checklist(rep78) tag(rep78 missing)

* Note that if you sort the output by 'make',
* the VW Diesel appears twice in the list; it has extreme values 
* for two of the variables we checked.  So if we went back to VW to 
* check the data, we would ask them about both variables.

* It can be quite powerful to sort the dataset and find ALL of the potential
* problems with a record or observation BEFORE going back to the source to
* check the data.

* Make a copy of this spreadsheet to preserve the original output as we want to make changes below.
copy "al_xl_demo.xlsx" "al_xl_demo_org.xlsx", replace

* We also want to make a copy so we can add replacement values to a file that has not been cleaned
copy "al_xl_demo.xlsx" "al_xl_demo_not_cleaned.xlsx", replace

********************************************************************************

* Once you have completed all potential assertions and BEFORE
* going back to the source to check the data, you can run the 
* assertlist_cleanup program to format columns and insert user friendly column headers.

* The first time through we will specify the required EXCEL option
* and optional NAME to preserve the original spreadsheet.
assertlist_cleanup, excel(al_xl_demo.xlsx) name(al_xl_demo_clean.xlsx)

* Next we will also specify the optional IDSORT option so that all sheets are
* sorted by the IDLIST provided in the original assertion. 
* This automatically does the sort by `make' mentioned above.
assertlist_cleanup, excel(al_xl_demo.xlsx) name(al_xl_demo_clean_and_sorted.xlsx) idsort

* Lastly we will only specify the required EXCEL option.
* This will overwrite the original EXCEL file with the cleaned up sheets.
* so we can show how assertlist_replace can be used on either EXCEL files. 
assertlist_cleanup, excel(al_xl_demo.xlsx)

********************************************************************************
* To show an example of assertlist_replace we will need to add some values to 
* the assertlist fix spreadsheets in the replace columns
* NOTE: These values are completely random and not all lines will contain 
* corrected values.

* If you are using Stata v15 you can simply run the changes below to obtain
* corrected values to use in the assertlist_replace demo.

* If you are running Stata v14 or earlier due to version limitations 
* you can either run the code below, then manually open up each spreadsheet 
* and SAVE the changes to populate the corrected values or populate your 
* own corrected values to each spreadsheet.

* You can also go into these excel files and make your own changes in the columns for replace values

foreach v in al_xl_demo_org al_xl_demo_clean al_xl_demo_not_cleaned {
	
	putexcel set "`v'.xlsx", modify sheet("want_to_correct_fix")	
	
	* Add replacement values for var 1
	putexcel I2 = 4, nformat(#)
	putexcel I3 = 2, nformat(#)
	putexcel I4 = 3, nformat(#)
	putexcel I5 = 4, nformat(#)
	putexcel I6 = 4, nformat(#)
	putexcel I7 = 3.5, nformat(#.#)
	
	* Adding replacement values for var 2 
	putexcel M8 = 190, nformat(#)
	putexcel close
	 
	 
	putexcel set "`v'.xlsx", modify sheet("several_tests_fix")
	* Add replacement values for var 1
	putexcel I2= 15900, nformat(#)
	putexcel I4 = 2, nformat(#)
	putexcel I5 = 3, nformat(#)
	putexcel I6 = 1, nformat(#)
	putexcel I7 = 4, nformat(#)
	putexcel I8 = 4, nformat(#)
	putexcel I9 = 2, nformat(#)
	putexcel I10 = 2, nformat(#)
	putexcel I11 = 1, nformat(#)
	putexcel I12 = 1, nformat(#)
	putexcel I13 = 3 , nformat(#)
	putexcel I14 = 4, nformat(#)
	putexcel I20 = 3, nformat(#) // conflict from other tab
	putexcel I21 = 2, nformat(#) // same as other tab
	putexcel I22 = 5, nformat(#) //conflict from other tab
	putexcel I23 = 1, nformat(#) //conflict from other tab
	putexcel I24 = 4, nformat(#) // same as other tab
	
	* Add replacement values for var 2
	putexcel M18 = 14, nformat(#)
	putexcel close
}

********************************************************************************
* Now we will run the assertlist_replace program
* This program can be ran on an excel straight from assertlist and assertlist_cleanup

* In first example we will run on a spreadsheet that does not have any corrections
* Message sent to screen and no .DO file created.
assertlist_replace, excel(al_xl_demo)

* In the next example we will only provide the required EXCEL option. All other
* options will be set to the default value or not populated
* .DO file saved as replacement_commands
assertlist_replace, excel(al_xl_demo_org)

* Third example we will use a spreadsheet that did not go through assertlist_cleanup and add a name for .DO file
assertlist_replace, excel(al_xl_demo_not_cleaned) dofile(al_xl_demo_replace_commands_non_cleaned)

* While this program can run on the most basic input it also 
* gives the user has the option to add details to help with documentation.
* Next we will add the DATE reviewed and name of REVIEWER, 
* original dataset name (DATASET1) and name for new dataset (DATASET2)
* This example also shows that the command can be used with both the original 
* assertlist and assertlist_cleanup excel files. 

assertlist_replace, excel(al_xl_demo_clean) dofile(al_xl_demo_clean_replace_commands) ///
	reviewer(NAME HERE) date(2020-03-20) dataset1(auto) dataset2(auto_replace)

* We can also add a comment at the top of the do file that provides additional 
* information for documentation purposes.
assertlist_replace, excel(al_xl_demo_clean) dofile(al_xl_demo_clean_replace_commands_with_comment) ///
	reviewer(NAME HERE) date(2020-03-20) dataset1(auto) dataset2(auto_replace) ///
	comments(These values were selected at random for example purposes and the changes should not be implemented in the auto dataset.)
	