* Examples of using assertlist, assertlist_cleanup and assertlist_replace on Stata's famous auto dataset
* Dale Rhoda
* October 16, 2018

sysuse auto, clear

tab rep78, m

* There are 5 missing values

assertlist !missing(rep78)

assertlist !missing(rep78), list(make rep78)

assertlist !missing(rep78), list(make foreign rep78)

* Start with an empty Excel files

capture erase al_xl_test.xlsx

* With no list option, we export the line numbers of rows that fail assertion

assertlist !missing(rep78), excel(al_xl_test.xlsx) sheet(test1)

* In this case we list the make and the value of rep78, and we include an informative tag

assertlist !missing(rep78), excel(al_xl_test.xlsx) sheet(test2) list(make rep78) tag(Missing value for rep78)

* Now we request extra spreadsheet columns to maybe make some corrections to the data.

assertlist !missing(rep78), excel(al_xl_test.xlsx) sheet(test3) fix idlist(make) checklist(rep78) tag(Missing value for rep78)

* Assertion that does not fail
* Note that the program makes an entry in the Assertlist_Summary sheet, but does 
* not make a new sheet named 'test4'

assertlist inlist(foreign,1,0), excel(al_xl_test.xlsx) sheet(test4) list(make foreign)

assertlist gear_ratio < 3.8, list(make gear_ratio)

assertlist gear_ratio < 3.8, excel(al_xl_test.xlsx) sheet(test5) fix idlist(make) checklist(gear_ratio) tag(Gear ratio higher than expected)

* Let's look at an assertion involving an 'if' statement 

scatter headroom length

assertlist headroom <= 3 if length <= 180, list(make headroom length)

* These two observations violate our expectations...so we might go and look up 
* the numbers again and possible fix the data. But we don't yet know whether we 
* will replace the value of headroom or of length, so we list BOTH in the 
* checklist option

assertlist headroom <= 3 if length <= 180, excel(al_xl_test.xlsx) sheet(test6) fix idlist(make) checklist(headroom length) tag(Headroom seems large for small length)

* Finally, here is an example of sending output from several assertions to a 
* single worksheet which can be used for data cleaning.

* I don't know much about cars...probably the values in the dataset are fine.
* I have inserted some assertions here that look for large values for some
* of these variables and writes out which cars had the large values.  We will
* pretend that a data manager might go back and check the large values to 
* see whether they are correct or mis-typed.

assertlist price < 15900, excel(al_xl_test.xlsx) sheet(several_tests) fix idlist(make) checklist(price) tag(Price seems very high)
assertlist mpg < 40, excel(al_xl_test.xlsx) sheet(several_tests) fix idlist(make) checklist(mpg) tag(MPG seems very high)
assertlist rep78 < 5 if !missing(rep78), excel(al_xl_test.xlsx) sheet(several_tests) fix idlist(make) checklist(rep78) tag(rep78 seems very high)
assertlist headroom < 5, excel(al_xl_test.xlsx) sheet(several_tests) fix idlist(make) checklist(headroom) tag(Head room seems very high)
assertlist trunk < 22, excel(al_xl_test.xlsx) sheet(several_tests) fix idlist(make) checklist(trunk) tag(Trunk space seems very high)

* Note that if you sort the output by 'make',
* the VW Diesel appears twice in the list; it has extreme values 
* for two of the variables we checked.  So if we went back to VW to 
* check the data, we would ask them about both variables.

* It can be quite powerful to sort the dataset and find ALL of the potential
* problems with a record or observation BEFORE going back to the source to
* check the data.

********************************************************************************

* Once you have completed all potential assertions and BEFORE
* going back to the source to check the data, you can run the 
* assertlist_cleanup program to format columns and insert user friendly column headers.

* The first time through we will specify the required EXCEL option
* and optional NAME to preserve the original spreadsheet.

assertlist_cleanup, excel(al_xl_test.xlsx) name(al_xl_test_clean.xlsx)


* Next we will also specify the optional IDSORT option so that all sheets are
* sorted by the IDLIST provided in the original assertion. 
* This automatically does the sort by `make' mentioned above.

assertlist_cleanup, excel(al_xl_test.xlsx) name(al_xl_test_clean_and_sorted.xlsx) idsort

* Lastly we will only specify the required EXCEL option.
* This will overwrite the original EXCEL file with the cleaned up sheets.
* But for this example, let`s make a copy of the original assertlist EXCEL file 
* so we can show how assertlist_replace can be used on either EXCEL files. 

copy "al_xl_test.xlsx" "al_xl_test_2.xlsx", replace

assertlist_cleanup, excel(al_xl_test_2.xlsx)
exit 99

********************************************************************************
* To show an example of assertlist_replace we will need to add some values to 
* the assertlist out spreadsheets in the replace  
* NOTE: These values are completely random and not all lines will contain 
* corrected values.

* However due to limitations you can either run the code below, 
* then manually open up each spreadsheet and SAVE the changes to 
* populate replace statements. If this is not done the assertlist_replace demo 
* will not produce any results.

* Or you can skip the code below and manually make changes to the excel files. 
foreach v in al_xl_test al_xl_test_clean {
	
	putexcel set "`v'.xlsx", modify sheet("test3_fix")	
	putexcel I2 = 4, nformat(#)
	putexcel I3 = 2
	putexcel I4 = 3
	putexcel I5 = 4
	putexcel I6 = 4
	putexcel close
	
	putexcel set "`v'.xlsx", modify sheet("test6_fix")
	putexcel I2 = 3
	putexcel N3 = 190
	putexcel close
	 
	putexcel set "`v'.xlsx", modify sheet("several_tests_fix")
	putexcel I2= 15900
	putexcel I4 = 2
	putexcel I5 = 3
	putexcel I6 = 1
	putexcel I7 = 4
	putexcel I8 = 4
	putexcel I9 = 2
	putexcel I10 = 2
	putexcel I11 = 1
	putexcel I12 = 1
	putexcel I13 = 3 
	putexcel I14 = 4
	putexcel close
}

* In the first example we will only provide the required EXCEL option. All other
* options will be set to the default value or not populated

assertlist_replace, excel(al_xl_test)

* Add a name for DOFILE

assertlist_replace, excel(al_xl_test) dofile(al_xl_text_replace_commands)

* The user has the ability to add details to help with documentation.
* Next we will add the DATE reviewed and name of REVIEWER, 
* original dataset name (DATASET1) and name for new dataset (DATASET2)
* This example also shows that the command can be used with both the original 
* assertlist and assertlist_cleanup excel files. 

assertlist_replace, excel(al_xl_test_clean) dofile(al_xl_text_clean_replace_commands) ///
	reviewer(NAME HERE) date(10-16-2018) dataset1(auto) dataset2(auto_replace)

* We can also add a comment at the top of the do file that provides additional 
* information for documentation purposes.

assertlist_replace, excel(al_xl_test_clean) dofile(al_xl_text_clean_replace_commands_with_comment) ///
	reviewer(NAME HERE) date(10-16-2018) dataset1(auto) dataset2(auto_replace) ///
	comments(These values were selected at random for example purposes and the changes should not be implemented in the auto dataset.)

