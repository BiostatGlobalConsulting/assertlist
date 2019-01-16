*! assertlist_cleanup version 1.10 - Biostat Global Consulting - 2019-01-16

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
* 2019-01-10	1.10	MK Trimner		Replaced excel formatting to include fmtid
*										removed putexcel txtwrap formatting and included in fmtids
*******************************************************************************
*
* Contact Dale Rhoda (Dale.Rhoda@biostatglobal.com) with comments & suggestions.
*

program define assertlist_cleanup_v15

	syntax  , EXCEL(string asis) [ NAME(string asis) IDSORT ]
	

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
		local passthrough 0
		forvalues b = 1/`f' {
			
			* Bring in the sheet
			capture import excel using "`excel'.xlsx", describe
			
			* Capture the sheet name			
			local sheet `=r(worksheet_`b')'
		
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
			
			* If it is a fix sheet, sort the variables by id
			if "`=strpos("`sheet'","fix")'"!="0" {
			
				* Grab the max number of vars checked
				qui {
					capture confirm var _al_num_var_checked
					if _rc==0 {
						tempvar num_var_checked_l
						destring _al_num_var_checked, gen(`num_var_checked_l')
						qui summarize `num_var_checked_l'
						local max `=r(max)'
						drop `num_var_checked_l'
					}
				}
				if "`idsort'"!="" assertlist_cleanup_idsort, excel(`excel') sheet(`sheet')  max(`max')
			}
	
			* Remove _al from var names
			local n 1 
			
			noi di as text "Renaming variables and formatting columns..."
			foreach v of varlist * {
			
				local criteria `criteria' M`n'2(int)
				
				* Rename all the variables
				assertlist_cleanup_rename, excel(`excel') sheet(`sheet') n(`n') ///
					max(`max') var(`v') passthrough(`passthrough')
				
				* Format the tabs
				*assertlist_cleanup_format, excel(`excel') sheet(`sheet') n(`n') ///
				*	m1(`m`n'1') m2(`m`n'2') type(`type') replace(`replace') 
				
				local ++n
			}
		* Format header row
		assertlist_cleanup_format_header, excel(`excel') sheet(`sheet') passthrough(`passthrough')
		
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

	syntax, EXCEL(string asis) SHEET(string asis) MAX(int)

	noi di as text "Sort sheet by ID Variables..."

	qui {
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
			
		* Create local with the idlist
		* Start at the 3rd word in `e' as the first two are 
		* check_sequence and num_var_checked
		local elist
		forvalues i = 3/`enum' {
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
			PASSTHROUGH(string asis)

	qui {

		local v `var'
		
		* Reset two locals that will be trigger column width formatting
		local type
		local replace
		
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
			
		if "``v''"=="check_sequence" 	local `v' Assertion Completed Sequence Number 
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
		if "``v''"=="num_var_checked"	local `v' Number of Variables Checked in Assertion
		
		if `max'!=0 {
			forvalues i = 1/`max' {
				if "``v''"=="var_`i'"			local `v' Name of Variable `i'  Checked in Assertion
				if "``v''"=="var_type_`i'"		{
					local `v' Value type of Variable `i'
					local type yes
				}
				if "``v''"=="original_var_`i'"	local `v' Current Value	of Variable `i'
				if "``v''"=="correct_var_`i'"	{
					local `v' Blank Space for User to Provide Correct Value of Variable `i' 
					local m`n'1 20
				}
				if "``v''"=="replace_var_`i'"	{
					local `v' Stata Code to Be Used to Replace Current Value with Correct Value for Variable `i'
					local replace yes
				}
			}
		}
		
		* also create local with max of variable name
		local m`n'2 =length("``v''")
							
		* Put the new variable name into excel file
		putexcel set "`excel'.xlsx", modify sheet("`sheet'") 

		mata: st_local("xlcolname", invtokens(numtobase26(``v'n')))
		putexcel `xlcolname'1 = "``v''"
		
		if `n'==1 local passthrough `m`n'2'
		else local passthrough `passthrough' `m`n'2'
		
	
		* Pass through the locals
		foreach v in m`n'1 m`n'2 type replace passthrough {
			c_local `v' ``v''
		}
	}
end

********************************************************************************
********************************************************************************
******							Format Excel Header						   *****
********************************************************************************
********************************************************************************

capture program drop assertlist_cleanup_format_header
program define assertlist_cleanup_format_header

	syntax , EXCEL(string asis) SHEET(string asis) PASSTHROUGH(string asis)

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
		firstrow allstring clear
		describe
						
		local m_v=`=r(k)'
		local r_v=`=r(N)'
		
		
		local highlight 0	
		local i 1
		foreach v of varlist * {
			tempvar `v'_l
			gen ``v'_l'=length(`v')
			summarize ``v'_l'
			local m`i'1=min(`=`r(max)'+1',25)
			local m`i'2=word("`passthrough'",`i')
			drop ``v'_l'
			
			if strpos("`v'","Blank") > 0 & `highlight' == 0 local highlight `i'
			local ++i
		}
			
		* Create fontid for bold that will be added when appropriate
		mata: bold = b.add_fontid()
		mata: b.fontid_set_font_bold(bold, "on")
				
		forvalues i = 1/`m_v' {
			* Create the header format ids
			mata format_header_`i' = b.add_fmtid()
			mata: b.set_fmtid(1,`i',format_header_`i')
			
			* Since this is row 1, make them shaded, bold and horizontal aligned
			mata: b.fmtid_set_fontid(format_header_`i', bold)
			mata: b.fmtid_set_fill_pattern(format_header_`i', "solid","lightgray")
			mata: b.fmtid_set_horizontal_align(format_header_`i', "left")
			mata: b.fmtid_set_text_wrap(format_header_`i', "on")
						
			* Set column width
			mata format_width_`i' = b.add_fmtid()
			mata: b.set_fmtid(2,`i',format_width_`i')
			
			* Set column width
			local width `=`m`i'1'+3'
			if `m`i'2'>`m`i'1' local width `=`m`i'1'+ 11'
			mata: b.fmtid_set_column_width(format_width_`i',`i',`i',`=min(30,`width')')
			mata: b.fmtid_set_text_wrap(format_width_`i', "on")
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
				mata: b.fmtid_set_text_wrap(format_highlight_`v', "on")
			
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
