*! assertlist_cleanup version 1.15 - Biostat Global Consulting - 2021-03-31

* This program can be used after assertlist to cleanup the column
* names and make them more user friendly

*******************************************************************************
* Change log
* 				Updated
*				version
* Date 			number 	Name			What Changed
* 2018-02-08	1.00	MK Trimner		Original Version
* 2018-04-10	1.01	MK Trimner		Added sorting option and broke into programs
* 2018-05-10	1.02	MK Trimner		Corrected typo
* 2018-06-06	1.03	MK Trimner		Added code to strip .xls or .xlsx extension
* 2018-09-13	1.04	MK Trimner		Added code to destring id variables when possible
*										Had to add code to create a temp string variable as well
*										in rename program
* 2018-09-27	1.05	MK Trimner		Added code to string .xls or .xlsx extension from NAME
* 2018-10-04	1.06	MK Trimner		Added a min of 30 for column width to prevent errors with 
*										column widths that are too long
* 										Also added txtwrap for entire sheet after all other formatting
*										is completed.
* 2018-10-10	1.07	MK Trimner		Corrected message
* 2018-10-24	1.08	Dale Rhoda		Use numtobase26() to pull the Excel column name we need
* 2018-11-21	1.09	MK Trimner		-Removed CONCATENATE formula (Replace command) and put
*										in replace program to speed up process.
*										- Added code to set column width to 0 if 
*										replace column or var type
* 2019-02-21	1.10	MK Trimner		Cleaned up formatting excel subprogram to remove code  
*										Also changed it to format each tab rather than each column at a time
*										Adjusted column width criteria so it takes into account the difference
*										between varname and value lengths rather than if name is greater
*										than values
* 2019-04-14	1.11	MK Trimner		Removed column for replace statement
*										and all code related
* 2020-04-09	1.12	MK Trimner		Added code to allow for LIST option with FIX spreadsheets
*										changed column width
*										Added new names for Assertlist Summary tab to show idlist checklist and list values
* 2020-08-13	1.13	MK Trimner		Added code to cleanup ID tab if populated
*										Changed syntax check for "fix" tabs
* 2020-09-07	1.14	MK Trimner		Added idsort to non-fix tabs as well
* 2020-09-09			MK Trimner		Removed Completed from sequence number	
* 2021-03-31	1.15	MK Trimner		Added noFormat option to allow Stata to run faster and avoid Excel formatting errors due to large size	
*******************************************************************************
*
* Contact Dale Rhoda (Dale.Rhoda@biostatglobal.com) with comments & suggestions.
*

program define assertlist_cleanup

	syntax  , EXCEL(string asis) [ NAME(string asis) IDSORT noFORMAT]

	noi di as text "Confirming excel file exists..."
	
	* If the user specified a .xls or .xlsx extension in NAME or EXCEL, strip it off here
	foreach v in excel name {
		if lower(substr("``v''",-4,.)) == ".xls"  ///
			local `v' `=substr("``v''",1,length("``v''")-4)'
		if lower(substr("``v''",-5,.)) == ".xlsx" ///
			local `v' `=substr("``v''",1,length("``v''")-5)'
	}
	
	* Make sure file provided exists
	capture confirm file "`excel'.xlsx"
	if _rc!=0 {
		* If file not found, display error and exit program
		noi di as error "Spreadsheet provided in macro EXCEL does not exist." ///
				" Current value provided was: `excel'"
				
		noi di as error "Exiting program..."
		exit 99
				
	}
	else {
		
		* Describe excel file to determine how many sheets are present
		capture import excel using "`excel'.xlsx", describe
		local f `=r(N_worksheet)'
			
		* If user requests a new file name, create copy and save as NAME
		if "`name'"!="" {
			noi di as text "Making copy of excel file named: `name'.xlsx..."
			copy "`excel'.xlsx" "`name'.xlsx", replace
			
			* Set excel local to new file name
			local excel `name'
		}


		* Go through each of the sheets
		forvalues b = 1/`f' {
			
			* Create a local that will hold the length of each header
			local passthrough 0
			local hide 0
			
			* Bring in the sheet
			capture import excel using "`excel'.xlsx", describe
			
			* Capture the sheet name			
			local sheet `=r(worksheet_`b')'
			
			if "`sheet'" == "List of IDs failed assertions" assertlist_cleanup_id_tab, excel(`excel') sheet(`sheet') `format'
			else {
		
				* Import file
				noi di as text "Importing excel sheet: `sheet'..."
				import excel "`excel'.xlsx", sheet("`sheet'") firstrow clear allstring
				
				* Grab column count
				qui describe
				local columns = r(k)
				
				* Create a local with the cell range for sheet
				local range `=r(range_`b')'
					
				* Set local for max number of vars checked
				local max 0
				local start 2
				* If it is a fix sheet, sort the variables by id
				if lower(substr("`sheet'",-4,.)) == "_fix" {
				
					* Grab the max number of vars checked
					qui {
						capture confirm var _al_num_var_checked
						if _rc==0 {
							tempvar num_var_checked_l
							destring _al_num_var_checked, gen(`num_var_checked_l')
							qui summarize `num_var_checked_l'
							local max `=r(max)'
							local start 3
							drop `num_var_checked_l'
						}
					}
				}
				
				if "`sheet'" != "Assertlist_Summary" & "`idsort'"!="" assertlist_cleanup_idsort, excel(`excel') sheet(`sheet') max(`max') start(`start') `format'
		
				* Remove _al from var names
				local n 1

				noi di as text "Renaming variables and formatting columns..."
				foreach v of varlist * {
					
					* Rename all the variables
					assertlist_cleanup_rename, excel(`excel') sheet(`sheet') n(`n') ///
						max(`max') var(`v') passthrough(`passthrough') hide(`hide') `format'
								
					local ++n
				}
				
			* Format header row for each tab
			if "`format'" == "" assertlist_cleanup_format_header, excel(`excel') sheet(`sheet') ///
				passthrough(`passthrough') hide(`hide')
			
			}
		}
	}

end


********************************************************************************
********************************************************************************
******							Assertlist Cleanup Sort					   *****
********************************************************************************
********************************************************************************
capture program drop assertlist_cleanup_idsort
program define assertlist_cleanup_idsort

	syntax, EXCEL(string asis) SHEET(string asis) MAX(int) START(int) [noFORMAT]

	noi di as text "Sort sheet by ID Variables..."

	qui {
		* Double check that IDlist provided is the same as previously used
		local e
		foreach v of varlist * {
			if strpos("`e'","_al_assertion_syntax")==0  {
				local e `e' `v' 
				capture destring `v', replace
			}
		}
				
		* Determine the number of words in previous IDlist
		* Need to subtract 1 as _al_assertion_syntax is included in list
		local enum = `= wordcount("`e'") - 1'
			
		* Create local with the idlist
		* Start at the 3rd word in `e' as the first two are 
		* check_sequence and num_var_checked
		local elist
		forvalues i = `start'/`enum' {
			local elist `elist' `=word("`e'",`i')'
		}
							
		* sort by the ids found in elist
		sort `elist'
		
		* Export the new sorted data
		export excel using "`excel'.xlsx", sheet("`sheet'") sheetmodify ///
					firstrow(var) nolabel datestring("%tdDD/Mon/CCYY")
				
	}
end

********************************************************************************
********************************************************************************
******						Rename Excel Variables 						   *****
********************************************************************************
********************************************************************************
capture program drop assertlist_cleanup_rename
program define assertlist_cleanup_rename

syntax  , EXCEL(string asis) SHEET(string asis) N(int) MAX(int) VAR(varlist) ///
			PASSTHROUGH(string asis) HIDE(string asis) [noFORMAT]
	qui {

		local v `var'
		
		* Reset two locals that will be trigger column width formatting
		local hide_var
		
		* Grab the max length for formatting
		tempvar `v'_l
		if substr("`: type `v''",1,3) != "str" tostring(`v'), replace
		gen ``v'_l'=length(`v')
		
		qui summarize ``v'_l'
		local m`n'1=`=r(max) + 1'

		drop ``v'_l'
					
		local `v' `=subinstr("`v'","_al_","",1)'
			
		* Grab the var name and placement for putexcel purpose
		local `v'n `n'
			
		if "``v''"=="check_sequence" 	local `v' Assertion Sequence Number 
		if "``v''"=="obs_number" 		local `v' Observation Number in Dataset 
		if "``v''"=="assertion_syntax"	local `v' Assertion Syntax That Failed
		if "``v''"=="tag" 				local `v' User Specified Additional Information
		if "``v''"=="total"				{
			local `v' Total Number of Observations Included in Assertion
			local m`n'1 17
		}	
		if "``v''"=="number_passed"		local `v' Number That Passed Assertion
		if "``v''"=="number_failed"		local `v' Number That Failed Assertion
		if "``v''"=="note1"				local `v' Note
		if "``v''"=="sheet"			local `v' Sheet Name That Contains Assertion Output
		if "``v''"=="idlist"		local `v' Variables Provided in IDLIST Option
		if "``v''"=="list"			local `v' Variables Provided in LIST Option
		if "``v''"=="checklist"		local `v' Variables Provided in CHECKLIST Option

		if "``v''"=="num_var_checked"	local `v' Number of Variables Checked in Assertion
				
		if `max'!=0 {
			forvalues i = 1/`max' {
				if "``v''"=="var_`i'"			local `v' Name of Variable `i'  Checked in Assertion
				if "``v''"=="var_type_`i'"		{
					local `v' Value type of Variable `i'
					local hide_var yes
				}
				if "``v''"=="original_var_`i'"	local `v' Current Value	of Variable `i'
				if "``v''"=="correct_var_`i'"	{
					local `v' Blank Space for User to Provide Correct Value of Variable `i' 
					local m`n'1 20
				}
			}
		}
		
		* also create local with max of variable name
		local m`n'2 =length("``v''")
							
		* Put the new variable name into excel file
		putexcel set "`excel'.xlsx", modify sheet("`sheet'") 

		mata: st_local("xlcolname", invtokens(numtobase26(``v'n')))
		if "`format'" == "" putexcel `xlcolname'1 = "``v''", txtwrap bold left fpattern("solid", "lightgray")
		else putexcel `xlcolname'1 = "``v''"

		if "`hide_var'"=="yes" local hide `hide' ``v'n' 
		
		*if `n'==1 local passthrough `m`n'2'
		local passthrough `passthrough' `m`n'2'
	
		* Pass through the locals
		foreach v in passthrough hide { 
			c_local `v' ``v''
		}
	}
end

********************************************************************************
********************************************************************************
******							Format Excel Tabs						   *****
********************************************************************************
********************************************************************************

capture program drop assertlist_cleanup_format_header
program define assertlist_cleanup_format_header

	syntax , EXCEL(string asis) SHEET(string asis) PASSTHROUGH(string asis) HIDE(string asis)
	
	* Format the width of each column
	* use mata to populate table formatting
	qui {
		mata: b = xl()
		mata: b.load_book("`excel'.xlsx")
		mata: b.set_mode("open")
		
		mata: b.set_sheet("`sheet'")
		
		* Determine the column widths
		noi import excel using "`excel'.xlsx", sheet("`sheet'") ///
		firstrow allstring clear
		describe
						
		local m_v=`=r(k)'
		local r_v=`=r(N)'
		
		local i 1
		foreach v of varlist * {
			tempvar `v'_l
			gen ``v'_l'=length(`v')
			summarize ``v'_l'
			local m`i'1=min(`=`r(max)'+1',25)
			local m`i'2=word("`passthrough'",`i')
			drop ``v'_l'
			
			local ++i
		}
		
		forvalues i = 1/`m_v' {
			* Set column width
			local width = max(`=`m`i'1'+3',10)
			if `m`i'2' - `m`i'1' > 5 local width `=`m`i'1'+ 11'
			if `m`i'2' - `m`i'1' > 15 local width `=`m`i'1'+ 14'
			mata: b.set_column_width(`i',`i',`=min(30,`width')')
		}
		
		foreach l in `=substr("`hide'",3,.)' {
			mata: b.set_column_width(`l',`l',0)
		}
		
		* Set the row height 
		mata: b.set_row_height(1,1,80)
		
		mata b.close_book()		
	}
	
end

********************************************************************************
********************************************************************************
******						Rename Excel Variables 						   *****
********************************************************************************
********************************************************************************
capture program drop assertlist_cleanup_id_tab
program define assertlist_cleanup_id_tab

syntax  , EXCEL(string asis) SHEET(string asis) [noFORMAT]

	qui {
	    
		* Import file
		noi di as text "Importing excel sheet: `sheet'..."
		import excel "`excel'.xlsx", sheet("`sheet'") firstrow clear allstring
	    		
		destring _al_number_assertions_failed, replace
										
		summarize _al_number_assertions_failed
		local max = r(max)
		
		foreach v of varlist* {
		    
			local `v' `v'
			if "`v'" == "_al_idlist" 					local _al_idlist 							List of Variables Used to Identify Line in Assertion
			if "`v'" == "_al_number_assertions_failed" 	local _al_number_assertions_failed 		Number of Assertions Line Failed
		}
		
		forvalues i = 1/`max' {
		    local abbreviation st
			if `=substr("`i'",-1,1)' == 2 local abbreviation nd
			if `=substr("`i'",-1,1)' == 3 local abbreviation rd
			if inlist("`=substr("`i'",-1,1)'","4","5","6","7","8","9","0") local abbreviation th
			if `i' >= 11 & `i' <= 19 local abbreviation th
		    local _al_assertion_details`i' Assertion Tag or Syntax for `i'`abbreviation' Failed Assertion
		}
		
		* Put the new variable name into excel file
		putexcel set "`excel'.xlsx", modify sheet("`sheet'") 
		
		local n 1
		foreach v of varlist* {
		    			
		   	mata: st_local("xlcolname", invtokens(numtobase26(`n')))
			if "`format'" == "" putexcel `xlcolname'1 = "``v''", txtwrap bold left fpattern("solid", "lightgray")
			else putexcel `xlcolname'1 = "``v''"
			local ++n
		}
	}
end