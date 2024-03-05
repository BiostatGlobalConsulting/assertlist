*! assertlist_export_all_ids version 1.03 - Biostat Global Consulting - 2021-03-31

* This program can be used after assertlist or assertlist_cleanup to grab the list of IDs
* that failed all assertions in a spreadsheet and export to single tab.

*******************************************************************************
* Change log
* 				Updated
*				version
* Date 			number 	Name			What Changed
* 2020-08-13	1.00	MK Trimner		Original program
* 2020-09-01	1.01	MK Trimner		Corrected placement of sheetcount & datacount locals
* 2020-09-07	1.02	MK Trimner		Updated to accomodate idlist added to non-fix excel tabs
* 2021-0-31		1.03	MK Trimner		Added format option to allow for a faster Stata run 
*										and to avoid excel formatting errors for large spreadsheets
* 2022-03-29	1.04	MK Trimner		Added _al_obs_number to all assertions so changed this code to sort by _al_obs_number instead of idlist
*										idlist is used as a secondary sorting method
*******************************************************************************
*
* Contact Dale Rhoda (Dale.Rhoda@biostatglobal.com) with comments & suggestions.
*
capture program drop assertlist_export_ids
program define assertlist_export_ids

	syntax  , EXCEL(string asis) [noFORMAT]

	qui {
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

			* Now lets create a single dataset with the cards that need reviewed
			capture import excel using "`excel'.xlsx", describe
			local f `=r(N_worksheet)'
			
			local sheetcount 0
			local datacount 0

			forvalues b = 1/`f' {
								
				* Bring in the sheet
				capture import excel using "`excel'.xlsx", describe
						
				* Capture the sheet name			
				local sheet `=r(worksheet_`b')'
				
				if "`sheet'" != "Assertlist_Summary" {
					
					local ++sheetcount
					
					* Import file
					noi di as text "Importing excel sheet: `sheet'..."
					import excel "`excel'.xlsx", sheet("`sheet'") firstrow clear allstring
					
					* Only keep the variables needed 
					local clean 0
					capture confirm var _al_check_sequence
					if _rc != 0 {
					    local clean 1
						capture rename AssertionSequenceNumber		 	_al_check_sequence
						capture rename UserSpecifiedAdditionalInform 	_al_tag
						capture rename ObservationNumberinDataset		_al_obs_number
						capture rename AssertionSyntaxThatFailed		_al_assertion_syntax
						
						capture drop NumberofVariablesCheckedinA
						
					}
					capture drop _al_num_var_checked
						
					* Create local with keep variables 
					local keepvars 
					local vcount 0
					foreach v of varlist* {
						if `vcount' == 0 local keepvars `keepvars' `v'
						if "`v'" == "_al_tag" local ++vcount 
					}
					keep `keepvars'
						
					tempfile data
					save `data', replace
					
					capture confirm file "`data'" 
					if _rc == 0 {
						use "`data'", clear
									
						* Create an idlist from the variables
						local idlist
						foreach v of varlist* {
							if !inlist("`v'","_al_tag", "_al_check_sequence", "_al_assertion_syntax","_al_obs_number") local idlist `idlist' `v'
						}
				
						gen _al_idlist = "`idlist'"	
						replace _al_idlist = "_al_obs_number" if missing(_al_idlist)
						duplicates drop
				
						local ++datacount 
						if `sheetcount' == 1  & `datacount' == 1 {
							tempfile assertion_ids_for_review
							save `assertion_ids_for_review', replace
						}
						else {
							append using `assertion_ids_for_review'
							duplicates drop
							save `assertion_ids_for_review', replace
						}
					}
					
				}
			}
		
			noi di as text "Create one dataset..."
			* replace tag value
			replace _al_tag = _al_check_sequence + " : " + _al_tag if !missing(_al_tag)
			replace _al_tag = _al_check_sequence + " : " + _al_assertion_syntax if missing(_al_tag)
			
			drop _al_assertion_syntax
			
			rename _al_tag _al_assertion_details
			
			destring _al_check_sequence, replace
			capture destring _al_obs_number, replace
			
			* Now we want to sort based on IDs
			local orderlist
			levelsof _al_obs_number, local(id)
			foreach i in `id' {
				levelsof _al_idlist if _al_obs_number == `i', local(list)
				local uniquelist
				foreach w in `list' {
					local wc = wordcount("`w'")
					forvalues n = 1/`wc' {
						local uniquelist `uniquelist' `=word("`w'",`n')'
					}
				}
				local list2 : list uniq uniquelist
				replace _al_idlist = `"`list2'"' if _al_obs_number == `i'
				local orderlist `orderlist' `uniquelist'
				
				foreach v in `list2' {
					if "`v'" != "_al_obs_number" {
						levelsof `v' if _al_obs_number == `i', local(list)
						local uniquelist
						foreach w in `list' {
							local wc = wordcount("`w'")
							forvalues n = 1/`wc' {
								local uniquelist `uniquelist' `=word("`w'",`n')'
							}
							local list2 : list uniq uniquelist
							replace `v' = `"`list2'"' if _al_obs_number == `i'

						}
					}
				}
				
			}
			
			local order2 : list uniq orderlist
			local order2 =subinstr("`order2'","_al_obs_number","",.)
			di "`order2'"
					
			sort _al_obs_number _al_check_sequence `order2', stable
			bysort _al_obs_number: gen n = _n
				
			drop _al_check_sequence
			reshape wide _al_assertion_details, i(_al_obs_number) j(n)
			
			order _al_obs_number _al_idlist `order2'

			* create count of the number of assertions failed
			gen _al_number_assertions_failed = 0
			label var _al_number_assertions_failed "Total Number of assertions that child failed"
			foreach v of varlist _al_assertion_details* {
				replace _al_number_assertions_failed = _al_number_assertions_failed + 1 if !missing(`v')
			}
			
			order _al_idlist 
			order _al_number_assertions_failed, before(_al_assertion_details1)
			sort _al_obs_number _al_idlist `idlist'
			compress
			save `assertion_ids_for_review', replace

			export excel using "`excel'.xlsx", firstrow (var) sheet("List of IDs failed assertions", replace) nolabel
		
			*******************************************************************************
			* If the Spreadsheet thas run through assertlist_cleanup 
			* tidy up the variable names
			if `clean' == 1 assertlist_export_ids_clean_up, excel(`excel') sheet(List of IDs failed assertions) `format'
			*******************************************************************************
			
			if "`format'" == "" {
				* Format the excel spreadsheet
				noi di as text "Format excel sheet..."
				describe
				local col `=r(k)'
				local row `=r(N)+1'

				local i 1
				foreach v of varlist * {
					local varname = length("`v'")
					tostring `v', replace
					tempvar `v'_l
					gen ``v'_l'=length(`v')
					summarize ``v'_l'
					local varstring =`=`r(max)'+5'
					local uselength `varstring'
					if `varname' > `varstring' local uselength `=`varname'+2'
					local maxlength 20
					if `varname' > 20 local maxlength `varname'
					local m`i'=min(`uselength',`maxlength')
					drop ``v'_l'
					local ++i
				}
				
				* Now format the excel
				mata: b = xl()
				mata: b.load_book("`excel'.xlsx")
				mata: b.set_mode("open")
				mata: b.set_sheet("List of IDs failed assertions")
				
				forvalues i = 1/`col' {
					mata: b.set_column_width(`i',`i', `m`i'')
					mata: b.set_text_wrap((2,`row'),`i',"on")
				}

				mata: b.set_fill_pattern(1,(1,`col'),"solid","lightgray")
				mata: b.set_font_bold(1,(1,`col'),"on")
				mata: b.set_horizontal_align(1,(1,`col'),"left")
				
				mata b.close_book()	
			}
		}
	}
end

********************************************************************************
********************************************************************************
******	Rename Excel Variables if Assertlit_cleanup has already ran		   *****
********************************************************************************
********************************************************************************
capture program drop assertlist_export_ids_clean_up
program define assertlist_export_ids_clean_up

	qui {
	    
		syntax,  EXCEL(string asis) SHEET(string asis) [noFORMAT]
	    
		capture destring _al_number_assertions_failed, replace
										
		summarize _al_number_assertions_failed
		local max = r(max)
		
		foreach v of varlist* {
			local `v' `v'
			if "`v'" == "_al_idlist" 					local _al_idlist 						List of Variables Used to Identify Line in Assertion
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
			else 	putexcel `xlcolname'1 = "``v''"
			local ++n
		}
	}
end