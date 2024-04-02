* Examples of using assertlist, assertlist_cleanup and assertlist_replace on Stata's famous auto dataset
* Dale Rhoda & Mary Kay Trimner
* March 31, 2021

* Make sure you cd to the location you want to run your test as this program does create output.

********************************************************************************
* Multiple excel and .do files are created during this demo.
* We want to first wipe out any that may already be existing in the current directory
* For this demo we want to erase any old excel files 
* Start with an empty Excel files
foreach f in al_xl_demo al_xl_demo_no_format al_xl_demo_not_cleaned al_xl_demo_clean al_xl_demo_clean_and_sorted ///
			al_xl_demo_org al_xl_demo_no_format_org al_xl_demo_no_format_not_cleaned al_xl_demo_clean_and_sorted_no_format ///
			al_xl_demo_clean_no_format al_xl_demo_with_id_tab {
	capture erase `f'.xlsx
}

* We also want to remove any old do files that are created later in this program
foreach f in replacement_commands al_xl_demo_replace_commands al_xl_demo_replace_commands_non_cleaned al_xl_demo_clean_replace_commands al_xl_demo_clean_replace_commands_with_comment {
	capture erase `f'.do
}

* Open Stata's auto dataset
sysuse auto, clear
gen id = _n
label var id "Line number"

* Create a string variable for testing purposes
gen test1 = ""
replace test1 = "Test_string" in 17
replace test1 = "Test_string" in 27
replace test1 = "Test_string" in 47
label var test1 "Variable created to use replace program with string value"

* Create another variable that we want to be missing so we can replace it
* as missing in our testing later on
gen test2 = .
replace test2 = 1 in 17
replace test2 = 22 in 43


********************************************************************************
* First we will take a look at rep78. This will show there are 5 missing values.
tab rep78, m

* Run assertlist to show the line number for each contradiction to the screen
assertlist !missing(rep78)

* Run assertlist to show the make for each contradiction to the screen
*assertlist !missing(rep78), idlist(make)

* Run the assertiona nd show both the line number and make and rep78 for all contradictions
assertlist !missing(rep78), list(make rep78)

* You can include as many variables as you would like in the list option.
assertlist !missing(rep78), list(price mpg rep78 headroom trunk foreign) idlist(id)

********************************************************************************
* If you would like to save the list of contradictions, use the excel option

* With no list option, the line numbers of rows that fail assertion are exported
assertlist !missing(rep78), excel(al_xl_demo.xlsx) sheet(do_not_want_to_correct) format

* With no list option, the line numbers of rows that fail assertion are exported
* But also do not specify format option. This may speed up the Stata run and avoid excel formatting errors if spreadsheet is too large
assertlist !missing(rep78), excel(al_xl_demo_no_format.xlsx) sheet(do_not_want_to_correct) 

* Now we want to add to this tab, but pass through additional variables
* To do this with the non-fix option the idlist must be the same
* Here we are defaulting to the original _al_ob_number
assertlist !missing(rep78), excel(al_xl_demo.xlsx) sheet(do_not_want_to_correct) list(make rep78) format

* Show without specifying format option
assertlist !missing(rep78), excel(al_xl_demo_no_format.xlsx) sheet(do_not_want_to_correct) list(make rep78) 

* In this case we list the make and the value of rep78, and we include an informative tag
* This will add the new variables to the far right of the tab
* And the tag will be included in both the Assertlist_Summary and do_not_want_to_correct tabs.
* This first run will ERROR because the idlist is the not the same
*assertlist !missing(rep78), excel(al_xl_demo.xlsx) sheet(do_not_want_to_correct) idlist(make rep78) tag(Missing value for rep78)

* but if we change the tab name, it will work
*assertlist !missing(rep78), excel(al_xl_demo.xlsx) sheet(do_not_want_to_correct2) list(make rep78) tag(Missing value for rep78) //id(id)

* Add "_fix" to the end of the SHEET name.
* this will ERROR out as this is not acceptable
* assertlist !missing(rep78), excel(al_xl_demo.xlsx) sheet(do_not_want_to_fix) list(make rep78) tag(Missing value for rep78)

* But if we add the FIX option.. the SHEET name can include FIX in it
assertlist !missing(rep78), excel(al_xl_demo.xlsx) sheet(do_not_want_to_fix) fix idlist(make price) checklist(rep78) tag(Missing value for rep78) format

* Show without specifying format option
assertlist !missing(rep78), excel(al_xl_demo_no_format.xlsx) sheet(do_not_want_to_fix) fix idlist(make price) checklist(rep78) tag(Missing value for rep78) 

* If we want to be able to go back and put in a corrected value, specify the FIX option.
* This creates extra spreadsheet columns for all variables provided in CHECKLIST.
* The sheet name will include "_fix" at the end.

* If the user populates these columns with the correct value they can use the ASSERTLIST_REPLACE program
* to read in these values and make the replacements.
assertlist !missing(rep78), excel(al_xl_demo.xlsx) sheet(want_to_correct) fix idlist(make) checklist(rep78) tag(Missing value for rep78) format
* Show without specifying format option
assertlist !missing(rep78), excel(al_xl_demo_no_format.xlsx) sheet(want_to_correct) fix idlist(make) checklist(rep78) tag(Missing value for rep78) 


* Lets run the first fix test above but add a second variable to IDLIST.
* This will ERROR out because the IDLIST provided is a different IDLIST then used in the previous line but has the same sheet name.
* assertlist !missing(rep78), excel(al_xl_demo.xlsx) sheet(want_to_correct) fix idlist(make price) checklist(rep78) tag(Missing value for rep78)

* Lets run the previous assertlist but add a LIST option to show more variables
* This will ERROR as the IDLIST and LIST combination must be the same for it to put on the same tab
* assertlist !missing(rep78), excel(al_xl_demo.xlsx) sheet(want_to_correct) fix idlist(make) checklist(rep78) tag(Missing value for rep78) list(price mpg)

* Rerun with a new tab
assertlist !missing(rep78), excel(al_xl_demo.xlsx) sheet(want_to_correct_with_list) fix idlist(make) checklist(rep78) tag(Missing value for rep78) list(price mpg) format
* Show without specifying format option
assertlist !missing(rep78), excel(al_xl_demo_no_format.xlsx) sheet(want_to_correct_with_list) fix idlist(make) checklist(rep78) tag(Missing value for rep78) list(price mpg)

* So we can rerun this by either changing the SHEETNAME or IDLIST
* In this case we will change the SHEETNAME
* The sheet will include "_fix" at the end
assertlist !missing(rep78), excel(al_xl_demo.xlsx) sheet(want_to_correct2) fix idlist(make price) checklist(rep78) tag(Missing value for rep78) format
* Show without specifying format option
assertlist !missing(rep78), excel(al_xl_demo_no_format.xlsx) sheet(want_to_correct2) fix idlist(make price) checklist(rep78) tag(Missing value for rep78) 


* If sheet is not provided, then the program will use the assertion number for the sheetname
assertlist !missing(rep78), excel(al_xl_demo.xlsx) list(make price rep78) tag(Missing value for rep78) format
* Show without specifying format option
assertlist !missing(rep78), excel(al_xl_demo_no_format.xlsx) list(make price rep78) tag(Missing value for rep78) 

* If sheet is not provided, then the program will use the assertion number for the sheetname
* When fix is specified, suffix _fix will be added
* We will also list gear ratio
assertlist !missing(rep78), excel(al_xl_demo.xlsx) idlist(make price) check(rep78) list(gear_ratio) tag(Missing value for rep78) fix format
* Show without specifying format option
assertlist !missing(rep78), excel(al_xl_demo_no_format.xlsx) idlist(make price) check(rep78) list(gear_ratio) tag(Missing value for rep78) fix
* Now we will rerun this with formatting to show what happens if you add formatting to an unformatted sheet
* the Summary tab will be formatted and so will tab 8_fix
//assertlist !missing(rep78), excel(al_xl_demo_no_format.xlsx) sheet(8) idlist(make price) check(rep78) list(gear_ratio) tag(Missing value for rep78) fix 


********************************************************************************
* Assertions where all lines pass

* If excel is not specified... nothing happens as all lines passed the assertion
assertlist gear_ratio < 4, list(make gear_ratio) format

* When Excel is specified
* Note that the program makes an entry in the Assertlist_Summary sheet, but does 
* not make a new sheet named 'no_fails' or no_fails_fix
assertlist inlist(foreign,1,0), excel(al_xl_demo.xlsx) sheet(no_fails) list(make foreign) format
assertlist gear_ratio < 4, excel(al_xl_demo.xlsx) sheet(no_fails) fix idlist(make) checklist(gear_ratio) list(foreign) tag(Gear ratio higher than expected) format
********************************************************************************
* Let's look at an assertion involving an 'if' statement 
assertlist headroom <= 3 if length <= 180, list(make headroom length)

* These two observations violate our expectations...so we might go and look up 
* the numbers again and possible fix the data. But we don't yet know whether we 
* will replace the value of headroom or of length, so we list BOTH in the 
* CHECKLIST option
* Note this will append to the existing want_to_correct_fix tab as the idlist is the same
assertlist headroom <= 3 if length <= 180, excel(al_xl_demo.xlsx) sheet(want_to_correct) fix idlist(make) checklist(headroom length) tag(Headroom seems large for small length) format

* If we only want to view the variable in the IF statement we would add it to the LIST option.
* This will ERROR out though, because we did not put it on a new sheet 
* assertlist headroom <= 3 if length <= 180, excel(al_xl_demo.xlsx) sheet(want_to_correct) fix idlist(make) list(length) checklist(headroom) tag(Headroom seems large for small length)

* Repeat with a new sheet
assertlist headroom <= 3 if length <= 180, excel(al_xl_demo.xlsx) sheet(want_to_correct3) fix idlist(make) list(length) checklist(headroom) tag(Headroom seems large for small length) format
* Finally, here is an example of sending output from several assertions to a 
* single worksheet which can be used for data cleaning.

* I don't know much about cars...probably the values in the dataset are fine.
* I have inserted some assertions here that look for large values for some
* of these variables and writes out which cars had the large values.  We will
* pretend that a data manager might go back and check the large values to 
* see whether they are correct or mis-typed.

assertlist price < 15900, excel(al_xl_demo.xlsx) sheet(several_tests) fix idlist(make) checklist(price) tag(Price seems very high) list(price rep78 gear_ratio) format
assertlist mpg < 40, excel(al_xl_demo.xlsx) sheet(several_tests) fix idlist(make) checklist(mpg) tag(MPG seems very high) list(price rep78 gear_ratio) format
assertlist rep78 < 5 if !missing(rep78), excel(al_xl_demo.xlsx) sheet(several_tests) fix idlist(make) checklist(rep78) tag(rep78 seems very high) list(price rep78 gear_ratio) format
assertlist headroom < 5, excel(al_xl_demo.xlsx) sheet(several_tests) fix idlist(make) checklist(headroom) tag(Head room seems very high) list(price rep78 gear_ratio) format
assertlist trunk < 22, excel(al_xl_demo.xlsx) sheet(several_tests) fix idlist(make) checklist(trunk) tag(Trunk space seems very high) list(price rep78 gear_ratio) format

* Add a test with a second check variable... This will create var_2* variables
assertlist trunk < 22, excel(al_xl_demo.xlsx) sheet(several_tests) fix idlist(make) checklist(trunk mpg) tag(Trunk space seems very high) list(price rep78 gear_ratio) format

* Add check for missing rep78 so we can add some conflicts
assertlist !missing(rep78), excel(al_xl_demo.xlsx) sheet(several_tests) fix idlist(make) checklist(rep78) tag(rep78 missing) list(price rep78 gear_ratio)

* Add check for test string that will be used to show how to make a string missing in replace program
assertlist test1 == "", excel(al_xl_demo.xlsx) sheet(several_tests) fix idlist(make) checklist(test1 test2) tag(Test1 should be missing) list(price rep78 gear_ratio)
assertlist test2 == . , excel(al_xl_demo.xlsx) sheet(several_tests) fix idlist(make) checklist(test2 test1) tag(Test2 should be missing) list(price rep78 gear_ratio)
* Note that if you sort the output by 'make',
* the VW Diesel appears twice in the list; it has extreme values 
* for two of the variables we checked.  So if we went back to VW to 
* check the data, we would ask them about both variables.

* It can be quite powerful to sort the dataset and find ALL of the potential
* problems with a record or observation BEFORE going back to the source to
* check the data.

* Make a copy of this spreadsheet to preserve the original output as we want to make changes below.
copy "al_xl_demo.xlsx" "al_xl_demo_org.xlsx", replace
copy "al_xl_demo_no_format.xlsx" "al_xl_demo_no_format_org.xlsx", replace

* We also want to make a copy so we can add replacement values to a file that has not been cleaned
copy "al_xl_demo.xlsx" "al_xl_demo_not_cleaned.xlsx", replace
copy "al_xl_demo_no_format.xlsx" "al_xl_demo_no_format_not_cleaned.xlsx", replace
********************************************************************************
********************************************************************************
* Now we will run the program to grab all the ids from each tab and put them in 1 tab
* Users can opt to create a new excel file or add it to the existing excel file.
* This will run the spreadsheet that has not been cleaned first.
assertlist_export_ids, excel(al_xl_demo) format
* Show without FORMAT option 
assertlist_export_ids, excel(al_xl_demo_no_format) 
 ********************************************************************************
* Once you have completed all potential assertions and BEFORE
* going back to the source to check the data, you can run the 
* assertlist_cleanup program to format columns and insert user friendly column headers.

* The first time through we will specify the required EXCEL option
* and optional NAME to preserve the original spreadsheet.
* This sheet has an ID tab
assertlist_cleanup, excel(al_xl_demo.xlsx) name(al_xl_demo_clean.xlsx) format
* Show without FORMAT option 
assertlist_cleanup, excel(al_xl_demo_no_format.xlsx) name(al_xl_demo_clean_no_format.xlsx) 

* Next we will also specify the optional IDSORT option so that all sheets are
* sorted by the IDLIST provided in the original assertion. 
* This automatically does the sort by `make' mentioned above.

assertlist_cleanup, excel(al_xl_demo_org.xlsx) name(al_xl_demo_clean_and_sorted.xlsx) idsort format
* Show without FORMAT option 
assertlist_cleanup, excel(al_xl_demo_no_format_org.xlsx) name(al_xl_demo_clean_and_sorted_no_format.xlsx) idsort 

* Lastly we will only specify the required EXCEL option.
* This will overwrite the original EXCEL file with the cleaned up sheets.
* so we can show how assertlist_replace can be used on either EXCEL files.
* This also shows how the ID tab will be cleaned up 
assertlist_cleanup, excel(al_xl_demo.xlsx) format
* Show without FORMAT option 
assertlist_cleanup, excel(al_xl_demo_no_format.xlsx) 

 ********************************************************************************
* Run the assertlist_export_ids on a cleaned dataset
* This will also add the cleaned up headers to the new ID tab
assertlist_export_ids, excel(al_xl_demo_clean_and_sorted.xlsx) format
* Show without FORMAT option 
assertlist_export_ids, excel(al_xl_demo_clean_and_sorted.xlsx) 

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
	putexcel J2 = 4, nformat(#) // Conflict with several_tests_fix tab ... Duplicate with want_to_correct_with_list_fix tab... Will show up in Conflict portion of .do file
	putexcel J3 = 2, nformat(#) // Conflict with several_tests_fix tab and want_to_correct_with_list_fix tab 
	putexcel J4 = 3, nformat(#) // Duplicate with several_tests_fix tab and want_to_correct_with_list_fix tab 
	putexcel J5 = 4, nformat(#) // Conflict from want_to_correct_with_list_fix tab 
	putexcel J6 = 4, nformat(#) // Duplicate from want_to_correct_with_list_fix tab
	putexcel J7 = 3.5, nformat(#.#)
	
	* Adding replacement values for var 2 
	putexcel N8 = 190, nformat(#)
		
	* Add a replacement variable to empty rows 
	* These will be ignored
	
	* This one is missing the id to make the replacement
	putexcel J10 = "!MISSING!" // Will not be completed due to lack of information
	putexcel G10 = "headroom"
	
	
	* this one is missing the variable to make the replacement
	putexcel J12 = 4, nformat(#) // Will not be completed due to lack of information
	putexcel D12 = "Subaru"
	
	* add a replacement variable to empty rows that has identifiers and variable
	* This will be processed as it has all the required information
	putexcel J14 = 4, nformat(#)
	putexcel D14 = "Dodge Diplomat"
	putexcel G14 = "headroom"

	putexcel close
	
	
	
	putexcel set "`v'.xlsx", modify sheet("want_to_correct_with_list_fix")	
	
	* Add replacement values for var 1
	putexcel L2 = 4, nformat(#) // Duplicate with want_to_correct_fix ... Conflict with several_tests_fix... Will show up in the Conflict portion of the .do file
	putexcel L3 = 5, nformat(#) // Conflict with want_to_correct_fix tab and several_tests_fix tab 
	putexcel L4 = 3, nformat(#) // Duplicate with several_tests_fix tab and want_to_correct_fix tab 
	putexcel L5 = 1, nformat(#) // Conflict from want_to_correct_fix tab 
	putexcel L6 = 4, nformat(#) // Duplicate from want_to_correct_fix tab
	
	* Add a replacement variable to empty rows 
	* These will be ignored because it is missing the Var name and ID
	putexcel L8 = 2, nformat(#)

	putexcel close

	putexcel set "`v'.xlsx", modify sheet("several_tests_fix")
	* Add replacement values for var 1
	
	* these first sets will be processed as is
	putexcel M2= 15900, nformat(#)
	putexcel M3= 35, nformat(#)
	
	putexcel M4 = 2, nformat(#)
	putexcel M5 = 3, nformat(#)
	putexcel M6 = 1, nformat(#)
	putexcel M7 = 4, nformat(#)
	putexcel M8 = 4, nformat(#)
	putexcel M9 = 2, nformat(#)
	putexcel M10 = 2, nformat(#)
	putexcel M11 = 1, nformat(#)
	putexcel M12 = 1, nformat(#)
	putexcel M13 = 3 , nformat(#)
	putexcel M14 = 4, nformat(#)
	
	
	putexcel M15 = 3, nformat(#) 
	
	
	putexcel M16 = 21, nformat(#) // 1 Conflict on same tab
	putexcel M17 = 15, nformat(#) // 1 Conflict on same tab
	putexcel M18 = 11, nformat(#) // 1 Conflict on same tab
	putexcel M19 = 14, nformat(#) // 1 Conflict on same tab
	
	* Now add replacement values for var 1
	putexcel M20 = "!MISSING!" // Conflict with want_to_correct_fix & want_to_correct_with_list_fix tabs
	putexcel M21 = "!MISSING!" // Conflict with want_to_correct_fix & want_to_correct_with_list_fix tabs
	putexcel M22 =  3, nformat(#) // Duplicate with want_to_correct_with_list_fix tab and want_to_correct_fix tab
	putexcel M24 = "!MISSING!" // Creates duplicate on same tab
	putexcel M25 = "!MISSING!"
	putexcel M26 = "!MISSING!" // Creates duplicate on same tab
	putexcel M27 = "!MISSING!" 
	
	* Add replacement values for var 2
	putexcel P18 = 14, nformat(#)
	putexcel Q24 = "!MISSING!" // Creates duplicate on same tab
	putexcel Q25 = "!MISSING!"
	putexcel Q26 = "!MISSING!" // Creates duplicate on same tab
	putexcel Q27 = "!MISSING!" 
	
	
	* Add a replacement variable to empty rows 
	* These will be ignored
	putexcel M28 = 2, nformat(#)
	putexcel Q29 = "!MISSING!" 

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
