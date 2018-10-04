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

{pmore} If the {help assertlist##fix:FIX} option is 
		specified, the {help assertlist##idlist:IDlist} and 
		{help assertlist##checklist:CHECKlist} options specify which
		variables are listed, and the {bf:LIST} option is ignored. {p_end}

{pmore} {bf:NOTE: If {help assertlist##list:LIST} and }
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

{pmore} {bf: EXCEL} must also be accompanied by the 
        {help assertlist##sheet:SHEET} option. {p_end}

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
{pmore} 

{marker sheet}
{pstd} {bf:SHEET} - Name of Excel worksheet {p_end}

{pmore} If {help assertlist##excel:EXCEL} is specified, {bf:SHEET} must 
        also be provided. It must be a valid Excel sheet name and it CANNOT be 
		Assertlist_Summary. 
		{p_end}

{pmore} If the {bf:sheet} already exists, the new output is appended to the existing sheet.{p_end}

{pmore} If the {help assertlist##fix:FIX} option is not specified, 
        {bf:SHEET} will be populated with the following: {p_end}
{pmore2} 1. {bf:_al_sequence_number}:	{it: Sequential counter for assertions whose output was directed to this {cmd:EXCEL} file.} {p_end}
{pmore2} 2. {bf:_al_assertion_that_failed} {it: Contains {cmd:exp} syntax.} {p_end}
{pmore2} 3. {bf:_al_tag}:		{it:String provided in {help assertlist##tag:TAG}, if any.} {p_end}
{pmore2} 4. {bf:var list}: 		{it:Variables provided in {help assertlist##list:LIST} option for all assertions on specified {help assertlist##sheet:SHEET}}. {p_end}

{marker tag}
{pstd} {bf:TAG} - user-specified string to list with the output (Often a short 
       description of what you tested and why.)
	   {p_end}

{pmore} {bf:NOTE: Do {it:NOT} put double quotes around tag text.} {p_end}

{marker fix}
{pstd} {bf: FIX} - Generates a worksheet with additional columns to 
       help data managers correct (or 'fix') errant data values. {p_end}

{pmore} {bf: Note: The program works best when FIX output goes to a different }
		{bf: sheet than LIST output. So when the user specifies the FIX option, }
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
{pmore2} 6. {bf:for each variable in {help assertlist##checklist:CHECKlist}:} {p_end}
{pmore3} a. {bf:_al_var_#}:		{it:Variable name} {p_end}
{pmore3} b. {bf:_al_var_type_#}:	{it:Variable type} {p_end}
{pmore3} c. {bf:_al_original_var_#}:	{it:Original variable value} {p_end}
{pmore3} d. {bf:_al_correct_var_#}:	{it:Blank cell highlighted yellow.  May be }
            {it: populated later by the data manager, with a correct value that }
			{it: should update the dataset's current errant value.} {p_end}
{pmore3} e. {bf:_al_replace_var_#}:	{it:Excel concatenate formula that will be }
            {it:used to create Stata syntax for a replace command. If the data }
			{it:manager fills in a correct value, the replace command will }
			{it:auto-populate in Excel and the syntax may be copied to a .do file }
			{it:for data cleaning.} (We have written a command to copy the replace statements into a .do file. See {help assertlist_replace}.) {p_end}
{marker idlist}
{pstd} {bf:IDlist} - List of variables that uniquely identify each observation. 
        These variables will be included in the replace syntax for corrections. 
		{p_end}

{pmore} {bf:IDlist} is required if {help assertlist##excel:EXCEL} and 
        {help assertlist##fix:FIX} options are specified.{p_end}
{pmore} {bf:IDlist} should only be provided if {help assertlist##excel:EXCEL} 
        and {help assertlist##fix:FIX} options are specified; it will otherwise be ignored. 
		{p_end}
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

Biostat Global Consulting created two additional companion programs to be ran after {cmd:assertlist}: 
{pstd} {help assertlist_cleanup} - Cleans up excel file generated by assertlist. {p_end}
{pstd} {help assertlist_replace} - Pulls replace statements with corrected values from {cmd:assertlist}/{cmd:assertlist_cleanup} spreadsheet and puts in .do file. {p_end}

{title:See Also}
{help assert}
{help assertlist_cleanup}
{help assertlist_replace}



