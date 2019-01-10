*! assertlist version 2.09 - Mary Kay Trimner & Dale Rhoda - 2019-01-10
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
* 2019-01-10	2.09	MK Trimner			Used fmtid for excel formatting
*******************************************************************************
*
* Contact Dale Rhoda (Dale.Rhoda@biostatglobal.com) with comments & suggestions.
*

program assertlist
	version 11.1
	syntax anything(name=assertion equalok everything) [, KEEP(varlist) ///
	       LIST(varlist) IDlist(varlist) CHECKlist(varlist) TAG(string) ///
		   EXCEL(string asis) SHEET(string asis) FIX]
	
	
	preserve
	
	* First save current file as a tempfile to be used later as 
	* We will be importing throughout the next few steps if excel file
	* Already exists
	qui {
			tempfile hold
			save "`hold'", replace
			
		 * This program will call several subprograms  
		 * The first will check all input options
		 noi check_options, keep(`keep') list(`list') ///
			   idlist(`idlist') checklist(`checklist') ///
			   excel(`excel') sheet(`sheet') `fix' hold(`hold')
			   
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
			noi write_xl_summary, assertion(`assertion') excel(`excel') ///
			hold(`hold') summaryexists(`summaryexists') sheet(`sheet')
		}
		
		* If there were lines that failed the assertion, complete the below steps
		use "`hold'", clear
		summarize _al_asrt 
		if `=r(min)'== 0 {
		
			* Trim down dataset to the vars needed
			noi trimdown, keep(`varkeep') hold(`hold')
			
			* If FIX is specified, create the fix page
			if "`fix'"!=""	///
				noi write_fix_sheet, excel(`excel') sheet(`sheet') ///
				check(`checklist') id(`idlist') sheetexists(`sheetexists') ///
				hold(`hold') row(`row') num(`num') orgvarlist(`orgvarlist')
				
			* If excel is not specific, display results
			* If EXCEL option is not specified, display results on screen
			if "`excel'"==""  {
				if "`keep'" != "_al_obs_number" {
					noi list `keep', table noobs
				}
				else {
					noi di ""
					noi di "Dataset row numbers that contradict the assertion:"
					noi di as text "`msg'"
					noi list `keep', table noheader noobs
				}
			}
			
			* If EXCEL is specified, but not FIX
			if "`excel'"!="" & "`fix'"=="" noi write_nofix_sheet, excel(`excel') ///
			sheet(`sheet') sheetexists(`sheetexists') row(`row') orgvarlist(`orgvarlist')
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
		* noi di as text "Completing syntax checks..."
			
		local exitflag 0
		
		* The list option is a synonym for the keep option; park
		* the contents of list inside keep before proceeding
		* Removing any duplicate values
		foreach v in `keep' `list' {
			local llist `llist' `v'
			local ullist  : list uniq llist
			local llist   : list sort ullist
		}
		
		local keep `llist'
				
		* If EXCEL is populated, make sure sheet is populated
		if "`excel'" != "" & "`sheet'" == "" {
			noi di as error "Assertlist error: You must specify the SHEET " ///
							"option with the EXCEL option."
			noi di as text "`msg'"
			local exitflag 1
		}
		
		* If FIX is populated, check required variables
		if "`fix'"!="" & ("`idlist'"=="" | "`checklist'"=="" | "`excel'"=="" | "`sheet'"=="") {
			noi di as error "Assertlist error: You must specify the " ///
							"IDLIST, CHECKLIST, EXCEL and SHEET options with the FIX option."
			noi di as text "`msg'"
							
			local exitflag 1
		}
				
		* If FIX is populated, check KEEP is empty
		if "`fix'"!="" & "`keep'"!="" {
			noi di as text "Assertlist warning: Ignoring KEEP and LIST " ///
							  "values as they cannot be used with FIX option."
			
			* Clear out keep values
			local keep
		}
				
		* Trim SHEET to 27 characters if needed, add fix suffix
		if "`fix'"!=""	{
			local sheet "`=substr("`sheet'",1,`=min(27,`=strlen("`sheet'")')')'_fix"
			* Remove any double __ from the name
			local sheet "`=subinstr("`sheet'","__","_",.)'"
		}
		
		* Trim SHEET to 31 characters if need
		local sheet  "`=substr("`sheet'",1,`=min(31,`=strlen("`sheet'")')')'"
		
			
		* Check that if FIX option is not set, CHECKLIST and IDLIST are empty
		if "`fix'"=="" & ("`checklist'"!="" | "`idlist'"!="") {
				noi di as text "Assertlist warning: Ignoring " ///
				"CHECKLIST and IDLIST values as they are not used "  ///
				"with the KEEP option; only  "     ///
				"used with the FIX option."
							
				local checklist 
				local idlist
		}
		
		* If FIX and KEEP and not set, set to tempvar name obs_number
		if "`keep'" == "" & "`fix'"=="" {
			capture confirm variable _al_obs_number
			if _rc==0 {
				noi di as error "Assertlist error: This dataset already " ///
								"contains a variable named _al_obs_number."
				noi di as error "The assertlist program would like to " ///
								"generate a variable with that name because " ///
								"you called assertlist without the KEEP option."
				noi di as error "Either use the KEEP option or rename " ///
								"variable _al_obs_number and rerun assertlist."
				noi di as text "`msg'"
				local exitflag 1
			}
			else {
				gen _al_obs_number = _n
				label variable _al_obs_number "Dataset row number"
				local keep _al_obs_number
				save "`hold'", replace
			}
		}
		
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
		
		* Create local with variables that will be created
		local varcheck tag check_sequence assertion_syntax `varlist_fix'
		
		* Create local of unique keep, idlist and checklist variables
		local llist
		foreach v in `keep' `idlist' `checklist' {
			local llist `llist' `v'
			local ullist  : list uniq llist
			local llist   : list sort ullist
		}
		
		local varkeep `llist'

		* Identify if generated var exists in kept variables 
		foreach v in `varkeep' {
			* Check to see if generated vars exist in vars that are kept
			* If they do, user will need to rename vars and program will exit.
			foreach l in `varcheck' {
				if "`v'"=="_al_`l'" & {
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
			
			if `=r(max)' > 1 noi di as text "Assertlist warning: Variables provided in IDLIST do " ///
			"not uniquely identify each row. The program will continue, but " ///
			"be aware that this could create undesirable consequences when replacing " ///
			"the values and we advise that you make the IDLIST unique."
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
															
				* Grab the list of variables used in previous IDlist
				* This list will initially include variables _al_check_sequence 
				* and _al_num_var_checked, these will be excluded from list below
				* When actual check occurs.
				if "`fix'"!="" {
					* Double check that IDlist provided is the same as previously used
					local e
					foreach v of varlist * {
						if strpos("`e'","_al_assertion_syntax")==0  {
							local e `e' `v' 
						}
					}
					
					* Determine the number of words in previous IDlist
					* Need to subtract 1 as _al_assertion_syntax is included in list
					local enum = `= wordcount("`e'") - 1'
					
					* Create local with the old idlist
					* Start at the 3rd word in `e' as the first two are 
					* check_sequence and num_var_checked
					local elist
					forvalues i = 3/`enum' {
						local elist `elist' `=word("`e'",`i')'
					}
									
					if "`idlist'"!="`elist'" {
						noi di as error "Assertlist error: IDLIST(`idlist') does not "
						noi di as error "match IDLIST(`elist')"
						noi di as error "previously used on SHEET `sheet'"
						noi di as error "Either change IDLIST to match or change SHEET and rerun."
						noi di as text "`msg'"
						local exitflag 1
					}		
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
				sheetexists checklist idlist orgvarlist {
				
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
			SUMMARYexists(int) SHEET(string asis)

	qui {
		* Write Summary tab...
		* noi di "Writing Summary Tab..."
			
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
			str2045 _al_note1 using "`results'"
		
		* Count how many passed and failed the logical statement
		* noi di as text "Counting # that passed & failed the assertion..."
		count if _al_asrt == 1
		local passed = r(N)
			
		count if _al_asrt == 0
		di r(N)
		local num_fail = r(N)
		
		* Determine if all observations passed the assertion
		if `num_fail' == 0 {
			noi di as text "All observations passed the assertion."
				
			post `handle' ($SEQUENCE) (`"`assertion'"') ("`=_al_tag'") ///
				(`=`passed' + `num_fail'') (`passed') (`num_fail') ///
				("All observations passed the assertion")	
		}
		else {
			if `num_fail' == 1 {
				noi di as text ///
				"`num_fail' observation failed assertion; see tab `sheet' for more details."
				
				post `handle' ($SEQUENCE) (`"`assertion'"') ("`=_al_tag'") ///
					(`=`passed' + `num_fail'') (`passed') (`num_fail') ///
				("`num_fail' observation failed assertion; see tab `sheet' for more details.")
			}
			
			if `num_fail'  > 1 {
				noi di as text ///
				"`num_fail' observations failed assertion; see tab `sheet' for more details."
				
				post `handle' ($SEQUENCE) (`"`assertion'"') ("`=_al_tag'") ///
					(`=`passed' + `num_fail'') (`passed') (`num_fail') ///
				("`num_fail' observations failed assertion; see tab `sheet' for more details.")
			}
		}		
		
		* Close postfile
		capture postclose `handle'	
		
		use "`results'", clear
		
		compress
		if `summaryexists'==1 export excel using "`excel'.xlsx", sheet("Assertlist_Summary") ///
			sheetmodify cell(A`=$SEQUENCE+1')  
		
		if `summaryexists'==0 export excel using "`excel'.xlsx", sheet("Assertlist_Summary") ///
						sheetreplace cell(A1) firstrow(variable)
						
		* Format Summary Page
		* noi di as text "Formatting Summary tab..."
		format_sheet, excel(`excel') sheet(Assertlist_Summary)
	}	
end

********************************************************************************
********************************************************************************
******							Trimdown Dataset 						   *****
********************************************************************************
********************************************************************************

capture program drop trimdown
program define trimdown

	syntax ,  KEEP(varlist) HOLD(string asis)
	
	qui {
		* Running syntax checks...
		* noi di "Trimming down dataset..."
		
		* noi di as text ///
		* "Dropping all observations that passed the assertion..."
		* Drop if passed the assertion
		* Note that if the observation failed, the assertion 
		* variable _al_asrt== 0
		drop if inlist(_al_asrt,1,.)

		* Only keep the variables needed for output
		keep `keep' _al_assertion_syntax _al_tag _al_check_sequence 
				
		* Put variables in order
		order _al_check_sequence _al_assertion_syntax _al_tag `keep'
		
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
		SHEETexists(int) HOLD(string asis) ROW(int) NUM(int) [ ORGVARLIST(string asis)]
		
		
	qui {
		* Create data for fix tab...
		* noi di "Creating Fix Tab..."	

		use "`hold'", clear
		
		* Save the var types to be used later on
		foreach v in `idlist' `checklist' {
			local `v' `: type `v'' //`=substr("`: type `v''",1,3)'
		}	

		* Create a var that counts how many vars need checked
		* These will be provided in the syntax through checklist
		gen _al_num_var_checked=`num'
						
		* Create new vars that will be used in the Excel spreadsheet
		* to show the old var value, correct value & replace statement
		/*
		noi di as text "Creating variables to act as placeholders " ///
						"for columns in Excel spreadsheet that will " ///
						"contain the original variables, correct " ///
						"variable values, & Excel concatenate formula" 
		*/		
		* Create 5 variables for each var in CHECKLIST
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
			
			gen _al_replace_var_`i'=""
											
			local p `p' _al_var_`i' _al_var_type_`i' _al_original_var_`i' ///
				_al_correct_var_`i' _al_replace_var_`i' 
			
			* Check to see if checklist var is part of idlist
			* If not, drop
			if strpos("`idlist'", "`=word("`checklist'",`i')'")==0 ///
				drop `=word("`checklist'",`i')'		
		}
			
		* Order variables
		* noi di "Ordering variables..."
		order _al_check_sequence _al_num_var_checked ///
			`idlist' _al_assertion_syntax _al_tag `p' 
		
		save "`hold'", replace
		
		* If the fix sheet exists...
		if `sheetexists'==1 {
			* Append new results to existing spreadsheet
			noi di as text "Appending results to pre-existing " ///
						   "`excel'.xlsx sheet(`sheet')." 
			
			export excel using "`excel'.xlsx", sheet("`sheet'") sheetmodify ///
				cell(A`row') datestring("%tdDD/Mon/CCYY") 
				
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
		* noi di "Formatting FIX tab..."
		* Identify which columns will be highlighted
		local hi `=`=wordcount("`idlist'")' + 8'
		
		* Format Fix Sheet
		format_sheet, excel(`excel') sheet(`sheet') highlight(`hi')
	}	
end		 

********************************************************************************
********************************************************************************
******						Write Excel No-Fix Tab						   *****
********************************************************************************
********************************************************************************

capture program drop write_nofix_sheet
program define write_nofix_sheet
	
	syntax, EXCEL(string asis) SHEET(string asis) SHEETexists(int) ROW(int) [ ORGVARLIST(string asis)]
	
	qui {
		* Create no fix tab...
		* noi di as text "Creating No-FIX tab..."
		
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
				sheetmodify cell(A`row') datestring("%tdDD/Mon/CCYY") 
				
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
		* noi di as text "Formatting No-FIX tab..."
		format_sheet, excel(`excel') sheet(`sheet') 
	}
end

********************************************************************************
********************************************************************************
******							Format Excel Sheet						   *****
********************************************************************************
********************************************************************************
* Format tabs
capture program drop format_sheet
program define format_sheet

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
		
		* Create a generic format for all cells
		mata format_all = b.add_fmtid()
		mata: b.set_fmtid((1,`r_v'),(1,`m_v'),format_all)
		
		* Create fontid for bold that will be added when appropriate
		mata: bold = b.add_fontid()
		mata: b.fontid_set_font_bold(bold, "on")
		
		mata format_header = b.add_fmtid()
		mata: b.set_fmtid(1,(1,`m_v'),format_header)
		* Since this is row 1, make them shaded, bold and horizontal aligned
		mata: b.fmtid_set_fontid(format_header, bold)
		mata: b.fmtid_set_fill_pattern(format_header, "solid","lightgray")
		mata: b.fmtid_set_horizontal_align(format_header, "left")
		
		forvalues i = 1/`m_v' {
			* Create the header format ids
			mata format_header_`i' = b.add_fmtid()
			mata: b.set_fmtid(`=`r_v'+1',`i',format_header_`i')
						
			* Set column width
			mata: b.fmtid_set_column_width(format_header_`i',`i',`i', `m`i'')
			
		}
					
		* Highlight the correct values yellow
		if "`highlight'"!="0" {
		
			* Determine which rows need highlighted to pass through
			local hi
			forvalues i = `highlight'(5)`m_v' {
				local hi `hi' `i'
			}

			foreach v in `hi' {
				* Create fmtid for highlighting
				mata format_highlight_`v' = b.add_fmtid()
				mata: b.set_fmtid((2,`r_v'),`v', format_highlight_`v')
				mata: b.fmtid_set_fill_pattern(format_highlight_`v', "solid","yellow")
			
				* Now set fmtid to hide columns not needed
				mata format_hide_`=`v'+1' = b.add_fmtid()
				mata: b.set_fmtid((1,`r_v'),`=`v'+1',format_hide_`=`v'+1')
				mata: b.fmtid_set_column_width(format_hide_`=`v'+1',`=`v'+1',`=`v'+1',0)

				mata format_hide_`=`v'-2' = b.add_fmtid()
				mata: b.set_fmtid((1,`r_v'),`=`v'-2',format_hide_`=`v'-2')
				mata: b.fmtid_set_column_width(format_hide_`=`v'-2',`=`v'-2',`=`v'-2',0)
			}
		}
		mata b.close_book()	
	}
end		

********************************************************************************
********************************************************************************
