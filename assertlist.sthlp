{smcl}
{* *! version 1.0 06Oct2017}{...}
{vieweralsosee "[D] assert" "mansection D assert"}{...}
{viewerjumpto "Syntax" "assertlist##syntax"}{...}
{viewerjumpto "Description" "assertlist##description"}{...}
{viewerjumpto "Remarks" "assertlist##remarks"}{...}
{title:Title}

{phang}
{bf:assertlist} {hline 2} List observations that contradict an assert command


{marker syntax}{...}
{title:Syntax}

{p 8 16 2}
{opt assertlist} {it:{help exp}} {ifin} [{cmd:,} 
          {it:{help assertlist##list:LIST}}(varlist) 
		  {it:{help assertlist##excel:EXCEL}}(string)
		  {it:{help assertlist##sheet:SHEET}}(string)
	      	  {it:{help assertlist##tag:TAG}}(string) 
		  {it:{help assertlist##fix:FIX}}
		  {it:{help assertlist##idlist:IDlist}}(varlist) 
		  {it:{help assertlist##checklist:CHECKlist}}(varlist)
		  {it:{help assertlist##noformat:noFORMAT}}
 ] {p_end}

{synoptline}
{p2colreset}{...}
{p 4 6 2}

{marker description}{...}

{title:Description}


{pstd} {cmd:assertlist} is a wrapper for the {cmd:assert} command.  By itself, 
       {cmd:assert} only lists HOW MANY dataset rows contradict the assertion.  
	   {cmd:assertlist} helps understand WHICH rows contradict the assertion, 
	   and HOW.
	   {p_end}  
	   
{pstd} This can be very useful when cleaning a dataset: find observations that
       do not meet one of your expectations and list variables with unexpected
	   values and those that uniquely identify the row so you can
	   correct the unexpected value, or change it to missing.
	   {p_end}	 
	   
{pstd} In the simplest case you may check an assertion without specifying any 
       options, and assertlist will show the row numbers of observations that 
	   contradict the assertion.  Sometimes row numbers alone are helpful.
	   {p_end}

{pstd} The command options a) specify variables whose contents should be 
       listed, b) add helpful tags to remind you what assertion was contradicted,
	   c) direct output to an Excel spreadsheet, and d) insert additional 
	   columns in the spreadsheet to facilitate fixing unexpected values.
	   {p_end}
	   
{hline}

{title:Required Input} 

{marker exp}
{pstd} {bf:exp} - Logical expression that resolves to either TRUE or FALSE for 
       each row of the dataset. All rows where the expression is FALSE will be 
	   displayed on the screen (default) or sent to an
	   {help assertlist##excel:EXCEL} spreadsheet. 
	   {p_end}

{title:Optional Input} 
	   	   
{dlgtab:Customize Output} 
{marker list}
{pstd} {bf:LIST} - varlist whose contents are displayed in the output window 
       for rows where {bf:exp} is false{p_end}

{pmore} If the {bf:LIST} option is specified with the {help assertlist##fix:FIX} option,
		these variables are listed in addition to the variables in the 
		{help assertlist##idlist:IDlist} and {help assertlist##checklist:CHECKlist}
		options.
		{p_end}

{pmore} {bf:NOTE: If {help assertlist##list:LIST}, {help assertlist##list:IDLIST} and }
        {bf:{help assertlist##fix:FIX} are not }
        {bf:specified, assertlist simply displays the row numbers of all lines }
		{bf:that contradict the assertion.} {p_end}

{dlgtab:Send Output to Excel} 
{marker excel} 

{pstd} {bf: EXCEL} - Name of the {bf:Excel} workbook that will hold the output. 
       {p_end}

{pmore} This can include just the file name (goes to current folder) or a folder 
        path and file name. Do {it:NOT} include double quotes around the path 
		and filename for output excel file. {p_end}

{pmore} If the output file does not exist, a new file is created. If it does 
        exist, {cmd:assertlist} will attempt to add new output to the file.
		It is common to send output from numerous different assertions to
		the same output file.
		{p_end}

{pmore} When you specify {bf: EXCEL}, the command will always make an entry in the 
        worksheet named `{help assertlist##assertlist_summary:Assertlist_Summary}' (This is a summary sheet with a 
		hard-wired name.) 
		{p_end}
		
{pmore} If there are 0 exceptions to {cmd:exp}, the command will NOT generate a 
        worksheet to hold exceptions, but it will always make an entry in 
		the summary tab. 
		{p_end}
		
{pmore} If there are 1+ exceptions to {cmd:exp}, the command will make an 
        entry in the summary tab AND write the exceptions to the requested 
		{help assertlist##sheet:SHEET}.
		{p_end}

{marker assertlist_summary}
{pstd} {bf:Assertlist_Summary} - Summary sheet that contains the following information for each assertion: {p_end}

{pmore2} 1. {bf:_al_sequence_number}:	{it: Sequential counter for assertions whose output was directed to this {cmd:EXCEL} file.} {p_end}
{pmore2} 2. {bf:_al_assertion_that_failed} {it: Contains {cmd:exp} syntax.} {p_end}
{pmore2} 3. {bf:_al_tag}:		{it:String provided in {help assertlist##tag:TAG}, if any.} {p_end}
{pmore2} 4. {bf:_al_total}: 		{it: Total number of observations {cmd:exp} was evaluated on.  (Depends on [if] and [in].)} {p_end}
{pmore2} 5. {bf:_al_number_passed}:	{it: Number of observations for which {cmd:exp} was TRUE.} {p_end}
{pmore2} 6. {bf:_al_number_failed}:	{it: Number of observations for which {cmd:exp} was FALSE.} {p_end}
{pmore2} 7. {bf:_al_note}		{it: Note regarding results.} {p_end}
{pmore2} 8. {bf:_al_sheet}		{it: Name of excel {help assertlist##sheet:SHEET} with results from assertion, if any.} {p_end}
{pmore2} 9. {bf:_al_idlist}		{it: Variables provided in {help assertlist##idlist:IDlist}, if any.} {p_end}
{pmore2} 10. {bf:_al_list}		{it: Variables provided in {help assertlist##list:LIST}, if any.} {p_end}
{pmore2} 11. {bf:_al_checklist}		{it: Variables provided in {help assertlist##checklist:CHECKlist}, if any.} {p_end}
{pmore} 

{marker sheet}
{pstd} {bf:SHEET} - Name of Excel worksheet {p_end}

{pmore} {bf:SHEET} is only an option when {help assertlist##excel:EXCEL} is specified.
		It must be a valid Excel sheet name and it CANNOT be 
		Assertlist_Summary. If {help assertlist##excel:EXCEL} option is specified and {bf:SHEET} is not provided, 
		the sheet will be the {bf:{_al_sequence_number}} or {bf:{_al_sequence_number}_fix} when the {help assertlist##fix:FIX} option is specified.
		{p_end}

{pmore} If the {bf:sheet} already exists, the new output is appended to the existing sheet.{p_end}

{pmore} {bf:Note: Do not include the string {it:"_fix"} at the end of any SHEET names as the program uses this to run certain steps.} {p_end}

{pmore} {bf:Note: If the {help assertlist##fix:FIX} option is specified, the SHEET can only have 1 set of {help assertlist##idlist:IDlist} and {help assertlist##list:LIST} variables.}
	{bf:If the user would like to complete an assertion with a different combination of {help assertlist##idlist:IDlist} and {help assertlist##list:LIST} variables}
	{bf:they will need to create a new SHEET.} {p_end}

{pmore} If the {help assertlist##fix:FIX} option is not specified, 
        {bf:SHEET} will be populated with the following: {p_end}
{pmore2} 1. {bf:_al_sequence_number}:	{it: Sequential counter for assertions whose output was directed to this {cmd:EXCEL} file.} {p_end}
{pmore2} 2. {bf:IDlist variables}: 	{it:One column for each of the variables in {help assertlist##idlist:IDlist}}. {p_end}
{pmore2} 3. {bf:_al_assertion_that_failed} {it: Contains {cmd:exp} syntax.} {p_end}
{pmore2} 4. {bf:_al_tag}:		{it:String provided in {help assertlist##tag:TAG}, if any.} {p_end}
{pmore2} 5. {bf:var list}: 		{it:Variables provided in {help assertlist##list:LIST} option for all assertions on specified {help assertlist##sheet:SHEET}}. {p_end}

{pmore} {bf:Note: If {help assertlist##fix:IDLIST} is not provided, assertlist sets IDLIST to row number.} {p_end}

{marker tag}
{pstd} {bf:TAG} - user-specified string to list with the output (Often a short 
       description of what you tested and why.)
	   {p_end}

{pmore} {bf:NOTE: Do {it:NOT} put double quotes around tag text.} {p_end}

{marker fix}
{pstd} {bf: FIX} - Generates a worksheet with additional columns to 
       help data managers correct (or 'fix') errant data values. {p_end}

{pmore} {bf: Note: The program works best when FIX output goes to a different }
		{bf: sheet than non-fix output. So when the user specifies the FIX option, }
		{bf: assertlist will }
        {bf: send output to a worksheet with the name }
		{bf: specified in the SHEET option PLUS the characters _fix.} 
		{p_end}

{pmore} When the user specifies the FIX option, the program will add these columns to the output: {p_end}
{pmore2} 1. {bf:_al_sequence_number}:	{it: Sequential counter for assertions whose output was directed to this {cmd:EXCEL} file.} {p_end}
{pmore2} 2. {bf:_al_num_var_checked}:	{it:Number of variables in the {help assertlist##checklist:CHECKlist}.} {p_end}
{pmore2} 3. {bf:IDlist variables}: 	{it:One column for each of the variables in {help assertlist##idlist:IDlist}}. {p_end}
{pmore2} 4. {bf:_al_assertion_that_failed} {it: Contains {cmd:exp} syntax.} {p_end}
{pmore2} 5. {bf:_al_tag}:		{it:String provided in {help assertlist##tag:TAG}, if any.} {p_end}
{pmore2} 6. {bf:LIST variables}:	{it: One column for each of the variables in {help assertlist##list:LIST}}.{p_end}
{pmore2} 7. {bf:for each variable in {help assertlist##checklist:CHECKlist}:} {p_end}
{pmore3} a. {bf:_al_var_#}:		{it:Variable name} {p_end}
{pmore3} b. {bf:_al_var_type_#}:	{it:Variable type} {p_end}
{pmore3} c. {bf:_al_original_var_#}:	{it:Original variable value} {p_end}
{pmore3} d. {bf:_al_correct_var_#}:	{it:Blank cell highlighted yellow.  May be }
            {it: populated later by the data manager, with a correct value that }
			{it: should update the dataset's current errant value.} {p_end}

{pmore} {bf:Note: If the data manager fills in a correct value and the {cmd:assertlist_replace} program is ran, the replace syntax is placed into a .do file for data cleaning purposes.} (See {help assertlist_replace}.) {p_end}
{marker idlist}
{pstd} {bf:IDlist} - List of variables that uniquely identify each observation. 
        These variables will be included in the replace syntax for corrections. 
		{p_end}

{pmore} {bf:IDlist} can be populated for any assertions. If {bf:IDLIST} is not specified, {cmd:assertlist} will set this option to the line number in dataset.{p_end}

{pmore} {bf:Note: It is best practice to use the same set of IDs across all assertions within the same excel file.}
	{bf: Be sure that ID variable or variables uniquely identify each row. If you do not have a unique ID, you can create one prior to running this program. } 
 
{marker checklist}
{pstd} {bf:CHECKlist} - List of variables used in {cmd:exp} that you may wish to correct later.  
        Every variable listed here will receive extra columns in the spreadsheet to facilitate corrections.{p_end}

{pmore} {bf:CHECKlist} must be provided if {help assertlist##excel:EXCEL} 
        and {help assertlist##fix:FIX} options are specified. {p_end}
{pmore} {bf:CHECKlist} should only be provided if 
        {help assertlist##excel:EXCEL} and 
		{help assertlist##fix:FIX} options are specified; it will 
		otherwise be ignored. {p_end}

{pmore} {bf: NOTE: CHECKlist should be all inclusive. If checking a date with }
        {bf:separate variables for month, day and year components, all three }
		{bf:components need to be provided if they might be corrected later. }
		{bf:It is sometimes convenient to hold a long list of variable names in a }
		{bf:local macro and list the macro in the IDlist or CHECKlist options.}
		{p_end}

{marker noformat}
{pstd} {bf:noFORMAT} - {cmd:assertlist} defaults to format columns with text, color and width options making the spreadsheet easy for the user to read. When {cmd:noformat} is specified all Excel formatting commands are ignored. 
This enables the user to run Stata faster and avoid potential Excel formatting errors due to large spreadsheets. {p_end}

{pmore2} {bf:NOTE: If a SHEET is formatted during a prior {help assertlist} run this formatting will not be undone.}
{bf:The same is true if a later run sends output to the same SHEET and does not specify {it:noformat}, assertlist will {it:ADD} formatting to that sheet.} {p_end}

{hline}

{pstd} Note that {cmd:assertlist} may fail due to: {p_end}
{pmore2} 1. Variables provided in the {help assertlist##list:LIST}, 
            {help assertlist##idlist:IDlist} or 
			{help assertlist##checklist:CHECKlist} options do not exist 
			in current dataset or the varnames are variables generated by 
			the {cmd:assertlist} program.{p_end}
{pmore2} 2. {help assertlist##exp:exp} is nonsensical and cannot be 
            evaluated or does not resolve to only TRUE or FALSE. {p_end}
{pmore2} 3. {help assertlist##fix:FIX} option specified but 
            {help assertlist##excel:EXCEL}, 
			{help assertlist##sheet:SHEET}, 
			{help assertlist##idlist:IDlist} and/or 
			{help assertlist##checklist:CHECKlist} are not provided. {p_end}
{pmore2} 4. {help assertlist##excel:EXCEL} option specified but 
            {help assertlist##sheet:SHEET} is not provided. {p_end}
{pmore2} 5. User specifies a {help assertlist##sheet:SHEET} named 
            "Assertlist_Summary". {p_end}

{title:Authors}
{p}

Mary Kay Trimner & Dale Rhoda, Biostat Global Consulting

Email {browse "mailto:Dale.Rhoda@biostatglobal.com":Dale.Rhoda@biostatglobal.com}

Biostat Global Consulting created three additional companion programs to be ran after {cmd:assertlist}: 
{pstd} {help assertlist_cleanup} - Cleans up excel file generated by assertlist. {p_end}
{pstd} {help assertlist_export_ids} - Provides a high level overview of results by creating a new excel tab within the assertion spreadsheet. 
				  This tab has a single row for each ID that fail 1 or more assertions with columns showing which assertions they failed. {p_end}
{pmore3}{bf: NOTE:To be run after {cmd:assertlist}/{cmd:assertlist_cleanup}}. {p_end}
{pstd} {help assertlist_replace} - Pulls all populated corrected variable values from fix worksheets within an assertlist spreadsheet and puts them in a .do file as replace statements. {p_end}
{pmore3}{bf: NOTE:To be run after {cmd:assertlist}/{cmd:assertlist_cleanup}}. {p_end}

{title:See Also}
{help assert}
{help assertlist_cleanup}
{help assertlist_export_ids}
{help assertlist_replace}




