*! assertlist version 2.19 - Mary Kay Trimner & Dale Rhoda - 2021-09-02
*******************************************************************************
* Change log
* 				Updated
*				version
* Date 			number 	Name				What Changed
* 2017-10-20	2.00	Mary Kay Trimner	Redesign & implement based on v1
*											and ideas for improvement
* 2017-08-26	2.01	Dale Rhoda			Added 'list' option
* 2017-11-21	2.01	Mary Kay Trimner	Changed variable name from 
*											`assertion_that_failed' to
*											`assertion_syntax'
*											also changed Excel Column list to local
*											`exlist'
* 2018-03-21	2.02	MK Trimner			Rearranged
* 2018-04-10	2.03	MK Trimner			Corrected no-fix sheet to check for previous variables
*											Formatted date variables to export in date format
*											Checked var type for CHECKLIST FIX variables
*											if they were in date format, 
*											change corresponding var_# format to match.
* 2018-09-13	2.04	MK Trimner			Changed var type to show full type 
*											Adjusted local in replace statement to
*											pull first 3 characters from the full var type
* 2018-09-27	2.05	MK Trimner			Added sheet name to Summary tab note
* 2018-10-04	2.06	MK Trimner			Changed the str## in post file to reflect max length of 2045
*											Corrected typos
* 2018-11-06	2.07	MK Trimner			Use numtobase26() to pull the Excel column name we need
* 2018-11-21	2.08	MK Trimner			Remove Replace Statements and put in Replace program
*											to speed up the process
*											Made the width of replace column and variable type
*											0 so they are hidden from the spreadsheet
*											allowing the user to see the relevant data
* 2019-02-19	2.09	MK Trimner			Added excel formatting to use fmtid when v15
*											Kept original format for v14
*											Set global at start of program with version number	
* 2019-04-17	2.10	MK Trimner			Removed column for replace statement	
*											Removed all comments with replace and any 
*											noi di that were commented out to clean up program	
* 2019-04-26	2.11	MK Trimner			Added wrap text in version 14 excel formatting	
*											Removed code to format and hide replace variables since these were removed		
* 2020-03-20	2.12	MK Trimner			Cleaned up comments	
* 2020-04-09	2.13	MK Trimner			Made changes to pass through list option for FIX spreadsheets	
* 2020-04-09	2.14	Dale Rhoda 	        Added some space around assertlist warning msgs			
* 2020-04-29	2.15	MK Trimner			Added code to create sheetname with $SEQUENCE number if not provided
*											Added nolabel option for exporting during all for consistency	
* 2020-08-04	2.16	MK Trimner			Added check to make sure sheet name does not end in fix
*											cleaned up syntax to see if is an assertlist "fix" tab	
* 2020-09-10	2.17	MK Trimner			Added idlist to nonfix excel option
*											and made other adjustments to accomodate this change
*											cleaned up error message that references keep to include list option
*											cleaned up comments on summry tab/screen
*											Removed comment to sheet_name in Assertlist_summary if all lines passed
*											Removed sections to export ids for all tabs as made a separate program
* 2021-03-31	2.18	MK Trimner			Added formatting option to prevent excel errors and speed up run
* 2021-09-02	2.19	MK Trimner			Removed SHEET message if all pass assertion
*******************************************************************************
*
* Contact Dale Rhoda (Dale.Rhoda@biostatglobal.com) with comments & suggestions.
*
program assertlist
	version 11.1
	syntax anything(name=assertion equalok everything) [, KEEP(varlist) ///
	       LIST(varlist) IDlist(varlist) CHECKlist(varlist) TAG(string) ///
		   EXCEL(string asis) SHEET(string asis) FIX noFORMAT]
	
	
	preserve
	
	* First save current file as a tempfile to be used later as 
	* We will be importing throughout the next few steps if excel file
	* Already exists
	qui {
		tempfile hold
		save "`hold'", replace
						
		* Grab the version of stata to be used for formatting
		if c(stata_version) < 15 global FORMATTING_VERSION 14
		else global FORMATTING_VERSION 15
			
		 * This program will call several subprograms  
		 * The first will check all input options
		 noi check_options, keep(`keep') list(`list') ///
			   idlist(`idlist') checklist(`checklist') ///
			   excel(`excel') sheet(`sheet') `fix' hold(`hold') `ids'
		
		* If everything passes the check
		* use `hold' file and generate assertion
		use "`hold'", clear
		capture gen _al_asrt = `assertion' 
		
		* Create variables to hold user input data: 
		* Assertion syntax and Tag (Blank if not specified) 
		gen _al_assertion_syntax = `"`assertion'"'
		gen _al_tag = "`tag'"
			
		* Create variable that shows check sequence
		gen _al_check_sequence=$SEQUENCE
		
		* If it is an invalid assertion, exit the program
		if _rc!=0 {
			noi di as error "Assertlist error: Invalid ASSERTION: `assertion'"
			noi di as error "Issue may be due to syntax error or variable " ///
							"in assertion does not exist in dataset."
			noi di as error "Correct issue and rerun."
			noi di as error "Exiting program"
			noi di as text "`msg'"
			exit 99
		}
		
		* Save changes to tempfile
		save "`hold'", replace
		
		* If EXCEL is specified write the excel summary
		* to show how many passed, failed and were included in assertion
		if "`excel'"!="" {
		    local passthroughoptions
		    if "`idlist'" != "" local passthroughoptions idlist(`idlist')
			if "`keep'" != "" local passthroughoptions `passthroughoptions' keep(`keep')
			if "`checklist'" != "" local passthroughoptions `passthroughoptions' checklist(`checklist')
			if "`sheetmessage'" != "" local passthroughoptions `passthroughoptions' sheetmsg(`sheetmessage')
			
			noi write_xl_summary, assertion(`assertion') excel(`excel') ///
			hold(`hold') summaryexists(`summaryexists') sheet(`sheet') `passthroughoptions' `format'
		}
		
		* If there were lines that failed the assertion, complete the below steps
		use "`hold'", clear
		summarize _al_asrt 
		if `=r(min)'== 0 {
		
			* Trim down dataset to the vars needed
			noi trimdown, keep(`varkeep') hold(`hold') idlist(`idlist')
			
			* If FIX is specified, create the fix page
			local fixpassthrough 
			if "`keep'"!="_al_obs_number" local fixpassthrough keep(`keep')
			if "`fix'"!=""	///
				noi write_fix_sheet, excel(`excel') sheet(`sheet') ///
				check(`checklist') id(`idlist') sheetexists(`sheetexists') ///
				hold(`hold') row(`row') num(`num') orgvarlist(`orgvarlist') `fixpassthrough' `format'
				
			* If excel is not specific, display results
			* If EXCEL option is not specified, display results on screen
			if "`excel'"==""  {
			    
				local message 
				local header
				if "`idlist'" == "_al_obs_number" local message "Dataset row numbers that contradict the assertion:" 
				if "`keep'" == "" & "`idlist'" == "_al_obs_number" local header noheader
					noi di ""
					noi di "`message'"
					noi di as text "`msg'"
					noi list `idlist' `keep', table noobs `header'
			}

			* If EXCEL is specified, but not FIX
			if "`excel'"!="" & "`fix'"=="" noi write_nofix_sheet, excel(`excel') ///
			sheet(`sheet') sheetexists(`sheetexists') row(`row') orgvarlist(`orgvarlist') `format'
			
		}
			
		* Bring back original dataset to Stata
		restore
	}
end	   
		
		
********************************************************************************
********************************************************************************
******						Check Input Options 						   *****
********************************************************************************
********************************************************************************

capture program drop check_options
program check_options

syntax [, KEEP(varlist) LIST(varlist) IDlist(varlist) CHECKlist(varlist) ///
		   EXCEL(string asis) SHEET(string asis) FIX HOLD(string asis) ]
	qui {	   
		* Running syntax checks...
		local exitflag 0
		
		* The list option is a synonym for the keep option; park
		* the contents of list inside keep before proceeding
		* Removing any duplicate values
		local llist
		foreach v in `keep' `list' {
			local llist `llist' `v'
			local ullist  : list uniq llist
			local llist   : list sort ullist
		}
		
		local keep `llist'
				
		* If EXCEL is populated, make sure sheet is populated
		c_local sheetmessage
		if "`excel'" != "" & "`sheet'" == "" c_local sheetmessage "Assertlist warning: Since option SHEET was not provided the SHEET will be populated with ASSERTION CHECK SEQUENCE number:" 
		noi di "`sheetmessage'"
		* If IDLIST is not set, create var name _al_obs_number
		if "`idlist'" == ""  {
			capture confirm variable _al_obs_number
			if _rc==0 {
				noi di as error "Assertlist error: This dataset already " ///
								"contains a variable named _al_obs_number."
				noi di as error "The assertlist program would like to " ///
								"generate a variable with that name because " ///
								"you called assertlist without the IDLIST option."
				noi di as error "Either use the IDLIST option or rename " ///
								"variable _al_obs_number and rerun assertlist."
				noi di as text "`msg'"
				local exitflag 1
			}
			else {
				gen _al_obs_number = _n
				label variable _al_obs_number "Dataset row number"
				local idlist _al_obs_number
				save "`hold'", replace
			}
		}
		
		* If FIX is populated, check required variables
		if "`fix'"!="" & ("`idlist'"=="" | "`checklist'"=="" | "`excel'"=="" ) {
			noi di as error "Assertlist error: You must specify the " ///
							"IDLIST, CHECKLIST, and EXCEL options with the FIX option."
			noi di as text "`msg'"
							
			local exitflag 1
		}
		
		* Check that user does not use FIX as ending of sheet name
		if lower(substr("`sheet'",-4,.)) == "_fix" & "`fix'" == "" {
		    noi di as error `"SHEET cannot end with the suffix "_fix" as this "' ///
							"is used by the assertlist program to identify tabs"
							
			noi di as text "`msg'"
			
			local exitflag 1
		}
				
		* Trim SHEET to 27 characters if needed, add fix suffix
		if "`fix'"!=""	{
			local sheet "`=substr("`sheet'",1,`=min(27,`=strlen("`sheet'")')')'_fix"
			* Remove any double __ from the name
			local sheet "`=subinstr("`sheet'","__","_",.)'"
		}
		
		* Trim SHEET to 31 characters if need
		local sheet  "`=substr("`sheet'",1,`=min(31,`=strlen("`sheet'")')')'"
		
		* Check that if FIX option is not set, CHECKLIST is empty
		if "`fix'"=="" & "`checklist'"!=""  {
				noi di ""
				noi di as input "Assertlist warning: Assertlist will ignore " ///
				"CHECKLIST values as they are only used with the FIX option."
				noi di as text "`msg'"
				local checklist 
		}
		
		* If FIX and KEEP are populated we want to set idlistpluskeep syntax to be used in messages
		* And create a variable with both lists
		local idandkeeplist `idlist' 
		if "`fix'" != "" local idandkeeplist `idandkeeplist' `keep'
		local idlistpluskeep IDLIST
		if "`keep'"!="" & "`fix'" != "" local idlistpluskeep IDLIST and LIST
	
		* Create local to check for variables that will be created
		if "`fix'"!="" {
			* For all variables provided in syntax...
			local varlist_fix
		
			forvalues i =1/`=wordcount("`checklist'")' {
				foreach v in var_`i' var_type_`i' ///
					original_var_`i' corrected_var_`i' {

					local varlist_fix `varlist_fix' `v'
				}
			}
		}
		
		* Add additional vars that will be created to th is local
		local varcheck tag check_sequence assertion_syntax `varlist_fix'
		
		* Create local of unique keep idlist and checklist variables
		foreach v in `keep' `idlist' `checklist' {
			local llist `llist' `v'
			local ullist  : list uniq llist
			local llist   : list sort ullist
		}
		
		local varkeep `llist'

		* Identify if vars generated in this program exist in kept variables 
		foreach v in `varkeep' {
			* Check to see if generated vars exist in vars that are kept
			* If they do, user will need to rename vars and program will exit.
			foreach l in `varcheck' {
				if "`v'"=="_al_`l'" {
					noi di as error "Assertlist error: Variable `v' is " ///
									"generated as a new variable in "    ///
									"assertlist program and exists in  " ///
									"current dataset."
					noi di as error "Rename variable `v' and rerun program."
					noi di as text "`msg'"
					local exitflag 1
				}
			}
		}
		
		* Check to see if IDLIST provided uniquely identifies respondent
		* If it does not, send warning to screen
		if "`idlist'" != "" {
			tempvar unique
			bysort `idlist': gen `unique'=_n
			summarize `unique'
			
			if `=r(max)' > 1 {
				noi di ""
				noi di as input "Assertlist warning: Variables provided in IDLIST do " ///
			"not uniquely identify each row. The program will continue, but " ///
			"be aware that this could create undesirable consequences when running " ///
			"assertlist_replace after fixing some of the values; " ///
			"we advise you to revise the IDLIST."
			noi di ""
			noi di as text "`msg'"
			}
		}
	
		* Set global SEQUENCE to 1 as default
		global SEQUENCE 1
		
		* Clean up excel file to remove extension
		* Check to see if excel sheet already exists
		* If it does, grab the sequence number
		* Grab the latest row number
		if "`excel'"!="" {
		
			* If the user specified a .xls or .xlsx extension, strip it off here
			if lower(substr("`excel'",-4,.)) == ".xls"  ///
				local excel `=substr("`excel'",1,length("`excel'")-4)'
			if lower(substr("`excel'",-5,.)) == ".xlsx" ///
				local excel `=substr("`excel'",1,length("`excel'")-5)'
					
			* Check to see if the Excel file for the log already exists
			capture import excel using "`excel'.xlsx", describe
			local f `=r(N_worksheet)'
					
			if "`f'"=="." local f 0 
			local summaryexists 0
			local sheetexists 0
			
			* If the EXCEL file exists, check to see if Assertlist_Summary 
			* and SHEET already exist as tabs; Two locals will be set and 
			* used later on to identify if the results need to be appended
			if `f'!=0 {
				forvalues sheetn=1/`f' {
					if "`=r(worksheet_`sheetn')'" == "Assertlist_Summary" 	local summaryexists 1
					if "`=r(worksheet_`sheetn')'" == "`sheet'" 				local sheetexists 1
				}		
			}
				
			* If the Assertlist_Summary tab already exists, grab the check_sequence value
			if `summaryexists'==1 {
				capture import excel using "`excel'.xlsx", ///
					sheet("Assertlist_Summary") firstrow clear 
				summarize _al_check_sequence
				global SEQUENCE `=r(max) + 1'
			}
			
			* check to see if SHEET needs to be changed to SEQUENCE NUMBER
			if "`sheet'" == "" 		local sheet $SEQUENCE
			if "`sheet'" == "_fix" 	local sheet ${SEQUENCE}_fix
			
			* If SHEET does not exist, set local ROW to 2
			if `sheetexists'==0 local row 2
			
			* set local to count how many words in checklist
			local num `=wordcount("`checklist'")'
			
			* If SHEET already exists grab data to know where to export to 
			* IDlist previously used
			if `sheetexists'==1 { 
				import excel using "`excel'.xlsx", sheet("`sheet'") firstrow clear
				
				* Check to see the number of variables that were previously checked
				describe, varlist
				
				* grab the row number to know where we need to export to
				* Add 2 to account for column names and where we want this placed
				local row `=r(N) + 2'
				
				* Grab the original varlist from sheet
				* This will help you confirm if the new vars match the old
				unab orgvarlist : _all 
															
				* Grab the list of variables used in previous IDlist and LIST if specified
				* This list will initially include variables _al_check_sequence 
				* and _al_num_var_checked, these will be excluded from list below
				* When actual check occurs.
				if "`fix'"!="" {
					* Double check that IDlist and LIST provided is the same as previously used
					local e
					local estart 3
					foreach v of varlist * {
						if strpos("`e'","_al_var_1")==0  {
							local e `e' `v' 
						}
					}
				}
				* Now do it for the non-fix sheets
				* Double check that IDlist and provided is the same as previously used
				else {
				    local estart 2
					local e
					foreach v of varlist * {
						if strpos("`e'","_al_assertion_syntax")==0 {
							local e `e' `v'
						}	
					}
				}
					
				* Determine the number of words in previous IDlist and LIST
				* Need to subtract 1 as _al_assertion_syntax is included in list
				local enum = `= wordcount("`e'") - 1'
				
				* Create local with the old IDLIST and LIST
				* Start at the 3rd word in `e' as the first two are 
				* check_sequence and num_var_checked
				local elist
				forvalues i = `estart'/`enum' {
					* Exclude _al_assertion_syntax and _al_tag
					if !inlist("`=word("`e'",`i')'","_al_assertion_syntax","_al_tag") local elist `elist' `=word("`e'",`i')'
				}
									
				if "`idandkeeplist'"!="`elist'" {
					noi di as error "Assertlist error: `idlistpluskeep'(`idandkeeplist') does not match `idlistpluskeep'(`elist') "
					noi di as error "previously used on SHEET `sheet'"
					noi di as error "Either change `idlistpluskeep' to match or change SHEET and rerun."
					noi di as text "`msg'"
					local exitflag 1
				}		
			}
		}	
		
		
		* If any of the above errors exist, exit program
		if `exitflag'==1 {
			noi di as error "Exiting program due to error(s) listed above"
			exit 99
		}
		
		else {
			* Pass through all the locals that will be used later on
			foreach v in keep varkeep row excel num varnum sheet summaryexists ///
				sheetexists checklist idlist orgvarlist idandkeeplist idlistpluskeep {
				
				c_local `v' ``v''
			}
		}
	}
end

********************************************************************************
********************************************************************************
******						Write Excel Summary Tab 					   *****
********************************************************************************
********************************************************************************

capture program drop write_xl_summary
program define write_xl_summary

	syntax, ASSERTION(string asis) EXCEL(string asis) HOLD(string asis) ///	
			SUMMARYexists(int) SHEET(string asis) ///
			[IDLIST(varlist) KEEP(varlist) CHECKLIST(varlist) SHEETMSG(string asis) noFORMAT] 

	qui {
		* Write Summary tab...
			
		* Bring in file
		use "`hold'", clear	
				
		* Create post file that will be used as log
		tempname handle
		tempfile results
			
		* Create a log file that will be used to capture how many passed 
		* and failed each assertion
		postfile `handle' _al_check_sequence ///
			str2045 _al_assertion_syntax ///
			str2045 _al_tag                   ///
			_al_total _al_number_passed _al_number_failed ///
			str2045(_al_note1 _al_sheet _al_idlist _al_list _al_checklist) using "`results'"
		
		* Count how many passed and failed the logical statement
		* noi di as text "Counting # that passed & failed the assertion..."
		count if _al_asrt == 1
		local passed = r(N)
			
		count if _al_asrt == 0
		local num_fail = r(N)
		
		* Determine if all observations passed the assertion
		if `num_fail' == 0 {
			noi di as text "All observations passed the assertion."
				
			post `handle' ($SEQUENCE) (`"`assertion'"') ("`=_al_tag'") ///
				(`=`passed' + `num_fail'') (`passed') (`num_fail') ///
				("All observations passed the assertion.") ///
				("") /// //("Tab `sheet' will not contain any output for this assertion") ///
				("`idlist'") ("`keep'") ("`checklist'")	
		}
		else {
			noi di "`sheetmsg'${SEQUENCE}"
						
			if `num_fail' == 1 {
				noi di as text ///
				"`num_fail' observation failed the assertion; see tab `sheet' for more details."
				
				post `handle' ($SEQUENCE) (`"`assertion'"') ("`=_al_tag'") ///
					(`=`passed' + `num_fail'') (`passed') (`num_fail') ///
				("`num_fail' observation failed the assertion. See appropriate tab for details.") ///
				("`sheet'") ("`idlist'") ("`keep'") ("`checklist'")	
			}
			
			if `num_fail'  > 1 {
				noi di as text ///
				"`num_fail' observations failed the assertion; see tab `sheet' for more details."
				
				post `handle' ($SEQUENCE) (`"`assertion'"') ("`=_al_tag'") ///
					(`=`passed' + `num_fail'') (`passed') (`num_fail') ///
				("`num_fail' observations failed the assertion. See appropriate tab for details") ("`sheet'") ///
				("`idlist'") ("`keep'") ("`checklist'")	
			}
		}		
		
		* Close postfile
		capture postclose `handle'	
		
		use "`results'", clear
		
		compress
		
		* Export results to Summary tab
		if `summaryexists'==1 export excel using "`excel'.xlsx", sheet("Assertlist_Summary") ///
			sheetmodify cell(A`=$SEQUENCE+1')  
		
		if `summaryexists'==0 export excel using "`excel'.xlsx", sheet("Assertlist_Summary") ///
						sheetreplace cell(A1) firstrow(variable)
						
		* Format Summary Page
		if "`format'" == "" format_sheet_v${FORMATTING_VERSION}, excel(`excel') sheet(Assertlist_Summary)
	}	
end

********************************************************************************
********************************************************************************
******							Trimdown Dataset 						   *****
********************************************************************************
********************************************************************************

capture program drop trimdown
program define trimdown

	syntax ,  KEEP(varlist) HOLD(string asis) IDlist(varlist)
	
	qui {
		* Drop if passed assertion...
		drop if inlist(_al_asrt,1,.)

		* Only keep the variables needed for output
		keep `keep' _al_assertion_syntax _al_tag _al_check_sequence 
				
		* Put variables in order
		order _al_check_sequence `idlist' _al_assertion_syntax _al_tag `keep'
		
		* Save tempfile with new changes
		save "`hold'", replace	
	}	
end

********************************************************************************
********************************************************************************
******						Write Excel Fix Tab 						   *****
********************************************************************************
********************************************************************************

capture program drop write_fix_sheet
program define write_fix_sheet

syntax, EXCEL(string asis) SHEET(string asis) IDlist(varlist) CHECKlist(varlist) ///
		SHEETexists(int) HOLD(string asis) ROW(int) NUM(int) ///
		[, KEEP(varlist) ORGVARLIST(string asis) noFORMAT]
		
		
	qui {
		* Create data for fix tab...
		use "`hold'", clear
		
		* Save the var types to be used later on
		foreach v in `idlist' `checklist' {
			local `v' `: type `v''
		}	

		* Create a var that counts how many vars need checked
		* These will be provided in the syntax through checklist
		gen _al_num_var_checked=`num'
						
		* Create new vars that will be used in the Excel spreadsheet
		* to show the old var value, correct value 
		* Create 4 variables for each var in CHECKLIST
		local p
		forvalues i =1/`num' {	
			
			gen _al_var_`i' = word("`checklist'",`i')
			
			gen _al_original_var_`i'=`=word("`checklist'",`i')'
			
			* Create variable with the vartype for the concatenate function
			gen _al_var_type_`i' = "``=word("`checklist'",`i')''"
			
			* Check to see if the vartype is a date function
			* If so, make the new var a date type
			if "`=substr("`:format `=word("`checklist'",`i')''",1,2)'"=="%t" {
				format %td _al_original_var_`i'
			}
			
			gen _al_correct_var_`i'=.
														
			local p `p' _al_var_`i' _al_var_type_`i' _al_original_var_`i' ///
				_al_correct_var_`i' 
			
			* Check to see if checklist var is part of idlist and keep combined local
			* If not, drop
			if strpos("`idlist' `keep'", "`=word("`checklist'",`i')'")==0 ///
				drop `=word("`checklist'",`i')'		
		}
			
		* Order variables
		order _al_check_sequence _al_num_var_checked ///
			`idlist' _al_assertion_syntax _al_tag `keep' `p' 
		
		save "`hold'", replace
		
		* If the fix sheet exists...
		if `sheetexists'==1 {
			* Append new results to existing spreadsheet
			noi di as text "Appending results to pre-existing " ///
						   "`excel'.xlsx sheet(`sheet')." 
			
			export excel using "`excel'.xlsx", sheet("`sheet'") sheetmodify ///
				cell(A`row') datestring("%tdDD/Mon/CCYY") nolabel
				
			* Export all the variable names 
			unab newvarlist: _all
			
			* If the new variable names do not match the old
			* Add the all the variable names to the spreadsheet
			if "`newvarlist'"!="`orgvarlist'" {
				local c 1
				foreach v in `newvarlist'  {
					putexcel set "`excel'.xlsx", modify sheet("`sheet'")
					
					mata: st_local("xlcolname", invtokens(numtobase26(`c')))
					putexcel `xlcolname'1 = ("`v'")
					
					local ++c
				}
			}
		}
		
		else {
			export excel using "`excel'.xlsx", sheet("`sheet'") ///
				sheetreplace firstrow(var) nolabel datestring("%tdDD/Mon/CCYY") 	 
		}
		
		* Format the spreadsheet

		* Identify which columns will be highlighted
		local hi `=`=wordcount("`idlist' `keep'")' + 8'
		
		* Format Fix Sheet
		if "`format'" == "" format_sheet_v${FORMATTING_VERSION}, excel(`excel') sheet(`sheet') highlight(`hi')
	}	
end		 

********************************************************************************
********************************************************************************
******						Write Excel No-Fix Tab						   *****
********************************************************************************
********************************************************************************

capture program drop write_nofix_sheet
program define write_nofix_sheet
	
	syntax, EXCEL(string asis) SHEET(string asis) SHEETexists(int) ROW(int) ///
			[ ORGVARLIST(string asis) noFORMAT]
	
	qui {
		* Create no fix tab...
		
		* if not fixing...
		* Export results to nonfix sheet
		if `sheetexists'==0 {
			no di as text "Results will be saved to a new sheet " ///
						  "in `excel'.xlsx file; sheet(`sheet')."
						  
			export excel using "`excel'.xlsx", sheet("`sheet'") ///
				sheetreplace firstrow(var) nolabel datestring("%tdDD/Mon/CCYY") 
			
		}
		else {
			no di as text "Results will be appended to " ///
						  "pre-existing `excel'.xlsx file, " ///
						  "sheet(`sheet')."
						  
			* Confirm all variables from sheet exist in dataset
			* If they do not, set them as missing
			foreach v in `orgvarlist' {
				capture confirm var `v'
				if _rc != 0 {
					gen `v' = .
				}
			}
			
			order `orgvarlist'
							
			* If the Excel sheet already exists append new results 
			* to existing spreadsheet
			export excel using "`excel'.xlsx", sheet("`sheet'") ///
				sheetmodify cell(A`row') datestring("%tdDD/Mon/CCYY") nolabel
				
			* Now do a putexcel to place the varnames
			* Create new local to be all varlist
			unab newvarlist: _all
			
			* Add the all the variable names to the spreadsheet
			* if the previous varlist and new varlist do not match
			if "`newvarlist'"!="`orgvarlist'" {
				local c 1
				foreach v in `newvarlist'  {
					putexcel set "`excel'.xlsx", modify sheet("`sheet'")
					
					mata: st_local("xlcolname", invtokens(numtobase26(`c')))
					putexcel `xlcolname'1 = ("`v'")
					
					local ++c
				}
			}
		}
		
		* Format tab
		if "`format'" == "" format_sheet_v${FORMATTING_VERSION}, excel(`excel') sheet(`sheet') 
	}
end

********************************************************************************
********************************************************************************
******							Format Excel Sheet						   *****
********************************************************************************
********************************************************************************
* Format tabs
capture program drop format_sheet_v14
program define format_sheet_v14

	syntax , EXCEL(string asis) SHEET(string asis) [ HIGHLIGHT(integer 0) ] 
	
	qui {
		* Pull in Excel sheet to format
		* Grab the Excel data: Row count
		* use mata to populate table formatting
		mata: b = xl()
		mata: b.load_book("`excel'.xlsx")
		mata: b.set_mode("open")
			
		* Determine the column widths
		noi import excel using "`excel'.xlsx", sheet("`sheet'") ///
		allstring clear
		describe
						
		local m_v=`=r(k)'
		local r_v=`=r(N)'
				
		local i 1
		foreach v of varlist * {
			tempvar `v'_l
			gen ``v'_l'=length(`v')
			summarize ``v'_l'
			local m`i'=min(`=`r(max)'+1',25)
			drop ``v'_l'
			local ++i
		}
		
		* use mata to populate table formatting
		mata: b = xl()
		mata: b.load_book("`excel'.xlsx")
		mata: b.set_mode("open")
			
		mata: b.set_sheet("`sheet'")
		
		* We want to wrap text for all content after header row
		* Create local that will do this if after the first row
		local tw
		forvalues i = 1/`m_v' {
			if `i' > 1 local tw , txtwrap 
			mata: b.set_column_width(`i',`i',`m`i'')`tw'
		}
			
		mata: b.set_fill_pattern(1,(1,`m_v'),"solid","lightgray")
		mata: b.set_font_bold(1,(1,`m_v'),"on")
		mata: b.set_horizontal_align(1,(1,`m_v'),"left")
			
		* format the column width on sheet page
		forvalues i = 1/`m_v' {
			mata: b.set_column_width(`i',`i',`m`i'')
		}
					
		* Highlight the correct values yellow
		if "`highlight'"!="0" {
		
			* Determine which rows need highlighted to pass through
			local hi
			forvalues i = `highlight'(4)`m_v' {
				local hi `hi' `i'
			}

			foreach v in `hi' {
				mata: b.set_fill_pattern((2,`r_v'),`v',"solid","yellow")
				mata: b.set_column_width(`=`v'-2',`=`v'-2',0)
			}
		}
		
		* If the sheet is Assertlist_Summary, we want to make 4 columns center aligned
		if "`sheet'" == "Assertlist_Summary" {
			foreach v in 1 4 5 6 {
				mata: b.set_horizontal_align((2,`r_v'),`v',"center")
			}
		}
		
		mata b.close_book()	
	}
end		

********************************************************************************
********************************************************************************
******						Format Excel Sheet	for v15					   *****
********************************************************************************
********************************************************************************
* Format tabs
capture program drop format_sheet_v15
program define format_sheet_v15

	syntax , EXCEL(string asis) SHEET(string asis) [ HIGHLIGHT(integer 0) ] 
	
	qui {
		* Pull in Excel sheet to format
		* Grab the Excel data: Row count
		* use mata to populate table formatting
		mata: b = xl()
		mata: b.load_book("`excel'.xlsx")
		mata: b.set_mode("open")
		
		mata: b.set_sheet("`sheet'")
			
		* Determine the column widths
		noi import excel using "`excel'.xlsx", sheet("`sheet'") ///
		allstring clear
		describe
						
		local m_v=`=r(k)'
		local r_v=`=r(N)'
				
		local i 1
		foreach v of varlist * {
			tempvar `v'_l
			gen ``v'_l'=length(`v')
			summarize ``v'_l'
			local m`i'=min(`=`r(max)'+1',25)
			drop ``v'_l'
			local ++i
		}
		
		* Create fontid for bold that will be added when appropriate
		mata: bold = b.add_fontid()
		mata: b.fontid_set_font_bold(bold, "on")
		
		* Add textwrap to all rows
		mata format_txtwrap = b.add_fmtid()
		mata: b.set_fmtid((2,`r_v'),(2,`m_v'),format_txtwrap)
		mata: b.fmtid_set_text_wrap(format_txtwrap, "on")
				
		forvalues i = 1/`m_v' {
			* Create the header format ids
			mata format_header_`i' = b.add_fmtid()
			mata: b.set_fmtid(1,`i',format_header_`i')
			
			* Since this is row 1, make them shaded, bold and horizontal aligned
			mata: b.fmtid_set_fontid(format_header_`i', bold)
			mata: b.fmtid_set_fill_pattern(format_header_`i', "solid","lightgray")
			mata: b.fmtid_set_horizontal_align(format_header_`i', "left")
						
			* Set column width
			mata format_width_`i' = b.add_fmtid()
			mata: b.set_fmtid(2,`i',format_width_`i')
			mata: b.fmtid_set_text_wrap(format_width_`i', "on")
			mata: b.fmtid_set_column_width(format_width_`i',`i',`i', `m`i'')
		}
					
		* Highlight the correct values yellow
		if "`highlight'"!="0" {
		
			* Determine which rows need highlighted to pass through
			local hi
			forvalues i = `highlight'(4)`m_v' {
				local hi `hi' `i'
			}

			foreach v in `hi' {
				* Create fmtid for highlighting
				mata format_highlight_`v' = b.add_fmtid()
				mata: b.set_fmtid((2,`r_v'),`v', format_highlight_`v')
				mata: b.fmtid_set_fill_pattern(format_highlight_`v', "solid","yellow")
			
				mata format_hide_`=`v'-2' = b.add_fmtid()
				mata: b.set_fmtid((1,`r_v'),`=`v'-2',format_hide_`=`v'-2')
				mata: b.fmtid_set_column_width(format_hide_`=`v'-2',`=`v'-2',`=`v'-2',0)
			}
		}
		
		* If the sheet is Assertlist_Summary, we want to make 4 columns center aligned
		if "`sheet'" == "Assertlist_Summary" {
			foreach v in 1 4 5 6 {
				mata: b.set_horizontal_align((2,`r_v'),`v',"center")
			}
		}

		mata b.close_book()	
	}
end		

********************************************************************************
********************************************************************************
