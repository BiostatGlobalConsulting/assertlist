*! assertlist_replace version 1.04 - Biostat Global Consulting - 2020-03-20

* This program can be used after assertlist and assertlist_cleanup to pull the 
* replace statements from the Excel file and put them in a .do file.

*******************************************************************************
* Change log
* 				Updated
*				version
* Date 			number 	Name			What Changed
* 2018-09-27	1.00	MK Trimner		Original Version
* 2018-10-17	1.01 	MK Trimner		Added code to only put output if replace
*										statements in spreadsheet...else, error is
*										sent to screen.
*										Moved the renaming of variables due to var
*										label to a local and outside of varlist loop
*										This prevents an error if multiple variables are 
*										used in CHECKLIST and renamed earlier
*										Also changed the check for Current value
*										to include more words in strpos so that replace
*										statement is not included
* 2018-11-21  	1.02	MK Trimner		- Added subprogram that populates replace 
*										statements in dataset and excel file...
*										this was removed from original Assertlist program
*										to speed up this process.
*										- Formatted columns in replace spreadsheet
*										to show replace commands
*										-Corrected split up of comments to remove duplicate first and last word
* 2019-04-17	1.03	MK Trimner		- Added some clean up steps: 
* 2019-05-01							delete the fix file at the end of program
*										capture postclose mkt before setting up postfile in case of error
*										Removed code to add replace statement to excel file
*										as well as formatting since it is not needed
*										Removed additional code that was not needed due to this change
* 2020-03-20	1.04	MK Trimner		Cleaned up comments
*******************************************************************************
*
* Contact Dale Rhoda (Dale.Rhoda@biostatglobal.com) with comments & suggestions.
*
capture program drop assertlist_replace
program define assertlist_replace

	syntax  , EXCEL(string asis) [DOfile(string asis) DATE(string asis) ///
								REVIEWER(string asis) COMMENTS(string asis) ///
								DATASET1(string asis) DATASET2(string asis)]
								
	* Grab the version of stata to be used for formatting
	if c(stata_version) < 15 global FORMATTING_VERSION 14
	else global FORMATTING_VERSION 15
	
	noi di as text "Confirming excel file exists..."
	
	* If the user specified a .xls or .xlsx extension, strip it off here
	if lower(substr("`excel'",-4,.)) == ".xls"  ///
			local excel `=substr("`excel'",1,length("`excel'")-4)'
	if lower(substr("`excel'",-5,.)) == ".xlsx" ///
			local excel `=substr("`excel'",1,length("`excel'")-5)'
			
	* Remove .do from DOFILE
	if lower(substr("`dofile'",-3,.)) == ".do"  ///
			local dofile `=substr("`dofile'",1,length("`dofile'")-3)'
			
	* Remove .dta from DATASET1 and DATASET2
	forvalues i = 1/2 {
		if lower(substr("`dataset`i''",-4,.)) == ".dta"  ///
			local dataset`i' `=substr("`dataset`i''",1,length("`dataset`i''")-4)'
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
		
		* If name not specified for dofile, set dofile name to default name
		if "`dofile'"=="" local dofile replacement_commands
	
		* Describe excel file to determine how many sheets are present
		capture import excel using "`excel'.xlsx", describe
		local f `=r(N_worksheet)'
				
		* Open post file 
		capture postclose mkt
		postfile mkt str100(sheet) float(sheetnum row varnum) str1000(tif trif assertion tag replacement varname) using fix, replace

		* Go through each of the sheets to determine if they are fix
		local cleanup 0
		forvalues b = 1/`f' {
			capture import excel using "`excel'.xlsx", describe
			local sheet `=r(worksheet_`b')'
		
			* If they are, pull the replace statements
			if "`=strpos("`sheet'","fix")'"!="0" {
			
				* Import file
				noi di as text "Importing excel sheet: `sheet'..."
				import excel "`excel'.xlsx", sheet("`sheet'") firstrow clear
				
				* Determine if there are any replace statements
				foreach v of varlist* {
					if strpos("`: var label `v''", "Blank Space") > 0 | strpos("`: var label `v''", "_al_correct")  {
						local cleanup 1
						qui count if !missing(`v')
						continue, break
					}
				}
						
				* Cleanup so we are only looking at the relevant information
				if `cleanup' == 1 assertlist_replace_cleanup, sheet(`sheet') sheetnum(`b') excel(`excel')
			}
		}
		
		postclose mkt
		
		* If there are lines left move on to the next steps
		if `cleanup'==1 & `=_N' > 0 {
			* Identify duplicates and conflicts
			assertlist_replace_conflict
				
			* Open .DO file and add opening comments
			assertlist_replace_open, excel(`excel') dofile(`dofile') comments(`comments') ///
				date(`date') reviewer(`reviewer') dataset1(`dataset1') dataset2(`dataset2')
		
			* Put the replace statements in .DO file
			assertlist_replace_commands, num(`c')
			
			* Add final save to .DO file
			file write replacement " save, replace" _n
			capture file close replacement
		}
		else noi di as error "No corrected values provided in spreadsheet."
		
		capture erase fix.dta
	}
	
end

********************************************************************************
********************************************************************************
******							Clean up data							   *****
********************************************************************************
********************************************************************************
capture program drop assertlist_replace_cleanup
program assertlist_replace_cleanup

	syntax , SHEET(string asis)	SHEETNUM(string asis) EXCEL(string asis)

	noi di as text "Adding `sheet' to full dataset..."

	qui {
		* Create locals to populate var number
		forvalues i = 1/4 {
			local `i' 1
		}
		
		* Create local with name of variables to drop and rename
		local droplist
		local renamelist
		
		foreach v of varlist * {
			if strpos("`: var label `v''","Name of Variable") > 0 {
				local renamelist `renamelist' rename `v' _al_var_`1' 
				local ++1
			}
				
			if strpos("`: var label `v''", "Blank Space") > 0 {
				local renamelist `renamelist' rename `v' _al_correct_var_`2' 
				local ++2
			}
			
			if strpos("`: var label `v''", "Current Value of ") > 0 {
				local renamelist `renamelist' rename `v' _al_original_var_`3'
				local droplist `droplist'  _al_original_var_`3'
				local ++3
			}
			
			if strpos("`: var label `v''", "Value type of ") > 0 {
				local renamelist `renamelist' rename `v' _al_var_type_`4'
				local droplist `droplist'  _al_var_type_`4'
				local ++4
			}
			
			if "`v'"=="UserSpecifiedAdditionalInform" rename `v' _al_tag
			if "`v'"=="AssertionSyntaxThatFailed" rename `v' _al_assertion_syntax
			if "`v'"=="AssertionCompletedSequenceNum" rename `v' _al_check_sequence
			if "`v'"=="NumberofVariablesCheckedinA" rename `v' _al_num_var_checked
		}
				
		* split up the rename local to execute each command
		local c `=wordcount("`renamelist'")'
		tokenize `renamelist'
		forvalues i = 1(3)`c' {
			``i'' ``=`i'+1'' ``=`i'+2''
		}
				
		* Create a local to grab the id variables
		local idlist
		foreach v of varlist * {
			if strpos("`v'","_al_") == 0 local idlist `idlist' `v' 
		}
		
		* Try to destring the idlist and other numeric variables
		foreach v in _al_check_sequence _al_num_var_checked `idlist' {
			destring `v', replace
		}
		
		if "`droplist'"=="" local droplist _al_original_* _al_var_type_*
				
			
		* Run subprogram to populate replace statements in dataset and excel file
		assertlist_pop_replace_statement, idlist(`idlist') //sheet(`sheet') excel(`excel') 
				
		* Drop variables not needed
		capture drop `droplist'
				
		* Drop if line does not contain a replace statement 
		gen num_vars=0
		qui summarize _al_num_var_checked
		forvalues i = 1/`=r(max)' {
			tostring _al_num_var_checked, replace
			replace num_vars = num_vars + 1 if !missing(_al_correct_var_`i')
		}
	
		drop if num_vars == 0
		drop num_vars
		
		* sort the local idlist to be in alphabetical order
		local idlist : list sort idlist
		
		* Create variable to show `if clause' 
		gen tif = ""
		forvalues i = 1/`=_N' {
			foreach id in `idlist' {
				if "`=substr("`:type `id''",1,3)'"=="str" ///
					replace tif = tif + " " + "`id'" + " " + "--" + strtrim(`id'[`i']) + "--" in `i'
				else replace tif = tif + " " + "`id'" + " " + string(`id'[`i'])	in `i'
			}
				
			* add a . if string value and missing
			replace tif = subinstr(tif,"----","--.--",.) in `i'
			replace tif = strtrim(tif) in `i'
			
			forvalues n = 1/`=_al_num_var_checked[`i']' {
				local tifvar `=_al_var_`n'[`i']' 
				local tifvarval `=_al_correct_var_`n'[`i']'
							
				if !inlist("`tifvarval'","",".") ///
					post mkt (`"`sheet'"') (`sheetnum') (`i') (`n') (`"`tifvar' `=tif[`i']'"') ///
					(`"`tifvar' `tifvarval' `=tif[`i']'"') (`"`=_al_assertion_syntax[`i']'"') ///
					(`"`=_al_tag[`i']'"') (`"`=_al_replace_var_`n'[`i']'"') (`"`=_al_var_`n'[`i']'"')
			}		
		}
	}
	
end	

********************************************************************************
********************************************************************************
******					Populate Replace Statement						   *****
********************************************************************************
********************************************************************************
capture program drop assertlist_pop_replace_statement
program assertlist_pop_replace_statement

	syntax , idlist(varlist) //sheet(string asis) excel(string asis) 
	
	qui {
	
		* Create a variable for the "if" poriton of the replace statement
		gen idlist="if "
		forvalues i = 1/`=_N' {
			foreach id in `idlist' {
				if "`=substr("`:type `id''",1,3)'"=="str" replace idlist = idlist + " " + "`id'" + " == " + `"""' + `id'[`i'] + `"""' + " &" in `i'
				else replace idlist = idlist + " " + "`id'" + " == " + string(`id'[`i']) + " &" in `i'
			}
			replace idlist = substr(idlist[`i'],1,length(idlist[`i'])-1) in `i'

		}
		replace idlist = strtrim(idlist)
		
		qui summarize _al_num_var_checked
		forvalues i = 1/`=r(max)' {
			
			gen _al_replace_var_`i'=""

			* Try to remove the line by line replace statements
			forvalues n = 1/`=_N' {
				if !missing(_al_correct_var_`i') in `n' {
					* Need to account for 4 different types of replace statements
					* Variable being replaced and Corrected Variable are both Strings
					if "`=substr(_al_var_type_`i'[`n'],1,3)'"=="str" & "`=substr("`:type _al_correct_var_`i''",1,3)'"=="str" ///
						& !inlist("`=_al_correct_var_`i'[`n']'","",".") {
						replace _al_replace_var_`i' = "replace `=_al_var_`i'[`n']'  = " ///
						+ `"""' + "`=_al_correct_var_`i'[`n']'" + `"""' + " " in `n'  
					}
					
					* Variable being replace is String but Corrected Variable is Numeric
					if "`=substr(_al_var_type_`i'[`n'],1,3)'"=="str" & "`=substr("`: type _al_correct_var_`i''",1,3)'"!="str" ///
						& !inlist("`=_al_correct_var_`i'[`n']'","",".") {
						replace _al_replace_var_`i' = "replace `=_al_var_`i'[`n']'  = " ///
						+ `"""' + string(`=_al_correct_var_`i'[`n']') + `"""' + " " in `n'  
					}
					
					* Variable being replaced and Corrected Variable are both Numeric
					if "`=substr(_al_var_type_`i'[`n'],1,3)'"!="str" & "`=substr("`: type _al_correct_var_`i''",1,3)'"!="str" ///
						& !inlist("`=_al_correct_var_`i'[`n']'","",".") {
						replace _al_replace_var_`i' = "replace `=_al_var_`i'[`n']' = " ///
						+ string(`=_al_correct_var_`i'[`n']') + " " in `n'  
					}
					
					* Variable being replaced is Numeric but Corrected Variable is String
					if "`=substr(_al_var_type_`i'[`n'],1,3)'"!="str" & "`=substr("`: type _al_correct_var_`i''",1,3)'"=="str" ///
						& !inlist("`=_al_correct_var_`i'[`n']'","",".") {
						replace _al_replace_var_`i' = "replace `=_al_var_`i'[`n']' = " ///
						+ "`=_al_correct_var_`i'[`n']'" + " " in `n'  
					}
					
					replace _al_replace_var_`i' = _al_replace_var_`i' + idlist if !missing(_al_replace_var_`i') in `n'
					
				}		
			}
		}
	}
end
********************************************************************************
********************************************************************************
******					Identify conflicting replace statement	 		  *****
********************************************************************************
********************************************************************************
capture program drop assertlist_replace_conflict
program define assertlist_replace_conflict

	noi di as text "Identifying conflicting and duplicate values..."

	qui  {
		* Bring in the `fix' dataset created in post file
		use "fix", clear
		compress
		
		* Create count of each `if clause' and replace statement
		bysort tif: gen tifn=_N
		bysort tif trif: gen trifn=_N
		
		* Create variable to show if there is a conflict between replace values
		gen conflict= tifn!=trifn
		
		* Create variable to show the number of conflicts
		bysort tif: gen num_conflict=_n if conflict==1
		
		* Create variable to show if duplicate
		gen duplicate=tifn==trifn if tifn > 1
		
		gen comment=""
		forvalues i = 1/`=_N' {
			replace comment = "* Duplicate: Replace statement shows up `=tifn[`i']' times in file." if duplicate==1  in `i'
			replace comment = "* Conflict: Replace statement shows up `=tifn[`i']' times, with 2+ replacement values." if conflict==1 	  in `i'
		}
		save "fix", replace
	}	
end

********************************************************************************
********************************************************************************
******						Open .DO file						 		  *****
********************************************************************************
********************************************************************************
capture program drop assertlist_replace_open
program define assertlist_replace_open

syntax , EXCEL(string asis) [ DOfile(string asis) COMMENTS(string asis) ///
							DATE(string asis) REVIEWER(string asis) ///
							DATASET1(string asis) DATASET2(string asis) ]
							
	noi di as text "Creating `dofile'.do file..."

	* Set local with file name
	local files replacement
								
	* Open .DO file
	file open replacement using `dofile'.do, text write replace
	file write replacement "* This program was automatically written by assertlist_replace command." _n
	file write replacement " " _n
	file write replacement "* This program will be used to run the replace commands from" _n


	qui {
		* Write header with excel information
		file write replacement "* `excel'.xlsx fix tabs." _n
		file write replacement "* .DO File created on $S_DATE" _n
		file write replacement " " _n
		
		* Add Date and Reviewer if provided
		if "`date'" != "" | "`reviewer'"!="" 	file write replacement "* These changes were reviewed: " _n
		if "`date'"!="" 						file write replacement "* On Date: `date'" _n
		if "`reviewer'"!=""						file write replacement "* By: `reviewer'" _n
		
		file write replacement " " _n
		
		* Now add code to open a dataset and save as new name if provided
		file write replacement " * Open original Dataset provided:" _n
		if "`dataset1'"!="" file write replacement `" use "`dataset1'", clear"' _n
		if "`dataset1'"=="" file write replacement `" use "ADD DATASET NAME HERE", clear"' _n
		file write replacement " " _n
		
		* Save file as new name
		file write replacement " * Save dataset with new name to preserve original values" _n
		
		* If a new name is not provided, set as default value
		if "`dataset2'" == "" local dataset2 dataset_with_replaced_values
		
		file write replacement `" save "`dataset2'", replace "' _n
		file write replacement " " _n
	
		* Add the comments if provided
		if "`comments'"!="" 	{
			local c `=wordcount("`comments'")'
		
			* split up the comments so that they are not too long in the .DO file
			
			tokenize "`comments'"
			forvalues i = 1(10)`c' {
					local comments`i'
				forvalues n = `i'/`=`i'+9' {
					local comments`i' `comments`i'' ``n''  
				}
				
				if `i'==1 	file write replacement "* Additional Comments: `comments1'" _n
				else 		file write replacement "* `comments`i''" _n
			}
		}
			
		* Add comments regarding conflicts at top of file
		qui count if num_conflict==1			
		local c `=r(N)'
		
		if `c'>0 {
			file write replacement " " _n
			file write replacement "********************************************************************************" _n
			file write replacement "********************************************************************************" _n
			file write replacement "* IMPORTANT NOTE TO USER: " _n 
			
			if `c' == 1 file write replacement "* There is `c' set of conflicting replace statements at the bottom of this .do file." _n
			else file write replacement "* There are `c' sets of conflicting replace statements at the bottom of this .do file." _n
			
			file write replacement "* Review each line and uncomment the replace statement with the correct value." _n
		}
			
		file write replacement " " _n
		
		* Pass through the number of conflicts
		c_local c `c'
	}
end

********************************************************************************
********************************************************************************
******					Put Replace statements in output				   *****
********************************************************************************
********************************************************************************
capture program drop assertlist_replace_commands
program assertlist_replace_commands

	syntax , NUM(string asis)

	noi di as text "Put replace statement in .DO file..."

	qui {
	
		* Create variable with dofile order
		gen dofile=string(sheetnum) + "_" + string(row) + "_" + string(varnum) 
		
		* Put in order in which it was received
		sort dofile, stable
		save, replace
		
		* First send out non-conflicting replace statements
		keep if conflict!=1
		
		* Put each replace statement out by sheet
		levelsof sheetnum, local(sh)
		foreach s in `sh' {
			preserve
			
			keep if sheetnum==`s'
			file write replacement "********************************************************************************" _n
			file write replacement "********************************************************************************" _n
			file write replacement " " _n
			file write replacement "* These replace statements are from sheet `=sheet[1]':" _n
			file write replacement " " _n
			
			* And By assertion
			levelsof assertion, local(g)
			tempfile data
			save "`data'", replace
			foreach v in `g' {
				use "`data'", clear
		
				keep if assertion=="`v'"

				local tag 
				if !inlist("`=tag[1]'","",".")  local tag / Tag: `=tag[1]'
			
				qui count 
				if r(N) > 0 {
					file write replacement "* Replacements made because:" _n
					file write replacement "* Failed assertion: `=assertion[1]' `tag'  " _n
				}

				forvalues i = 1/`=_N' {
					if "`=comment[`i']'"!="" file write replacement `"`=comment[`i']'"' _n
					file write replacement `"`=replacement[`i']'"' _n
					file write replacement " " _n	
				}
			}
			restore
		}
		
		* Now write out all the conflicts
		if `num'>=1 {
			use "fix", clear
			keep if conflict==1
			file write replacement "********************************************************************************" _n
			file write replacement "********************************************************************************" _n
			file write replacement "********************************************************************************" _n
			file write replacement "* This section contains conflicting replace statements." _n
			file write replacement "* Review each line and uncomment the statement with the correct value." _n
			
			* Create variable to show which conflict group they are a part of
			egen cgroup=group(tif conflict) if conflict==1
		
			* Put out conflicting replace statements by `if clause' and varname
			forvalues i = 1/`num' {
				preserve
				keep if conflict==1 & cgroup==`i'			
				file write replacement "********************************************************************************" _n
				file write replacement "* Conflict #`i' - `=tifn[1]' Lines " _n
				forvalues n = 1/`=tifn[1]' {
					file write replacement "* Line `n' " _n
					file write replacement `"* Sheet: `=sheet[`n']'     Row:`=row[`n']'     Variable Number:`=varnum[`n']'     Varname: `=varname[`n']'"' _n
					file write replacement `"* Assertion: `=assertion[`n']' 	Tag: `=tag[`n']'	"' _n
					file write replacement `"* `=replacement[`n']'"' _n
					file write replacement " " _n
				}
				file write replacement "********************************************************************************" _n
				file write replacement " " _n
				file write replacement " " _n
				restore
			}
		}
	}	
end
