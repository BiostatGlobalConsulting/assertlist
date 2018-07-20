* Examples of using assertlist on Stata's famous auto dataset
* Dale Rhoda
* November 29, 2017

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

* Note that if you sort the output by 'make' (after doing the find-and-replace 
* trick) that the VW Diesel appears twice in the list; it has extreme values 
* for two of the variables we checked.  So if we went back to VW to 
* check the data, we would ask them about both variables.

* It can be quite powerful to sort the dataset and find ALL of the potential
* problems with a record or observation BEFORE going back to the source to
* check the data.
