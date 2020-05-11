{smcl}
{* *! version 1.0 06Oct2017}{...}
{vieweralsosee "[D] assert" "mansection D assert"}{...}
{viewerjumpto "Syntax" "assertlist_replace##syntax"}{...}
{viewerjumpto "Description" "assertlist_replace##description"}{...}
{viewerjumpto "Remarks" "assertlist_replace##remarks"}{...}
{title:Title}

{phang}
{bf:assertlist_replace} {hline 2} Pulls all populated corrected variable values from {it:fix} worksheets within an {help assertlist} spreadsheet and puts them in a .do file as replace statements.

{marker syntax}{...}
{title:Syntax}
{p 8 16 2}
{opt assertlist_replace}{cmd:,}   {it:{help assertlist_replace##excel:EXCEL}}(string) [ {it:{help assertlist_replace##dofile:DOFILE}}(string)
{it:{help assertlist_replace##date:DATE}}(string) {it:{help assertlist_replace##reviewer:REVIEWER}}(string) 
{it:{help assertlist_replace##comments:COMMENTS}}(string) {it:{help assertlist_replace##dataset1:DATASET1}}(string)
{it:{help assertlist_replace##dataset2:DATASET2}}(string)]
{p_end}

{synoptline}
{p2colreset}{...}
{p 4 6 2}

{marker description}{...}
{title:Description}

{pstd} {cmd:assertlist_replace} is a companion for the {help assertlist} or {help assertlist_cleanup} commands. {cmd:assertlist} lists observations that contradict an 
	assert command and provides details around WHICH rows failed the assertion, and HOW. {cmd:assertlist_cleanup} takes the excel output from {cmd:assertlist} and adds user friendly column titles and formatting.{p_end}

{pstd} {cmd: assertlist_replace} is to be used after the user has reviewed each failed assertion and added the appropriate values to the {it:correct} columns in the 
	{cmd:assertlist} or {cmd: assertlist_cleanup} excel file. This program then puts each populated value in a replace statement from all {help assertlist##fix:FIX} tabs 
	directly into a .do file. Comments are added to show the sheetname, failed assertion and tag for each replace statement.
	{p_end}

{pstd} {cmd: assertlist_replace} also looks across all tabs for duplicate and conflicting values based on the {help assertlist##idlist:IDLIST} provided in the original assertlist command. 
	Duplicate statements are noted above the replace command so if a value needs to be changed the user knows how many other lines will also need to be updated. 
	Conflicting statements are commented out and put at the bottom of the .do file for review. The user will need to review each line and select the appropriate value by removing the {bf:*} before the replace statement.
	{p_end} 
{pstd} {cmd:assertlist_replace} can only be used if the {help assertlist##excel:EXCEL} and {help assertlist##fix:FIX} options were specified while running {cmd:assertlist}. 
	If the output file does not exist, the program will exit immediately. 
	{p_end}
	   
{hline}

{title:Required Input} 

{marker excel}
{pstd} {bf: EXCEL} - Name of the {cmd:Excel} workbook that holds the {cmd:assertlist}/{cmd:asserlist_cleanup} output with replace values. {p_end}
{pmore} {it:*See {help assertlist_replace##note:NOTE}  for additional information regarding {cmd:EXCEL}.}

{title:Optional Input} 
{marker dofile}
{pstd} {bf:DOFILE} - Name for .do file that will contain all the replace statements. {p_end}
{pmore}	This option allows the user to specify the name of the .do file. If {cmd:DOfile} is not provided, the default name is {bf:replacement_commands}. 
(You do not need to include the characters .do in the {cmd:DOfile} option.){p_end}
{pmore} {it:*See {help assertlist_replace##note:NOTE} for additional information regarding {cmd:DOfile}.} {p_end}
{marker date}
{pstd} {bf:DATE} - User can optionally specify the date when the review occurred, and this program will include it in a comment at the top of the .do file for documentation purposes. {p_end}
{marker reviewer}
{pstd} {bf:REVIEWER} - Name of person(s) who reviewed failed assertions and added correct values to spreadsheet. Included at the top of the .do file for documentation purposes. {p_end}
{marker comments}
{pstd} {bf:COMMENTS} - Additional notes user would like added to the top of the .do file.{p_end}
{marker dataset1}
{pstd} {bf:DATASET1} - Name of the original dataset used with {cmd:assertlist} command. This is auto-populated into a {cmd:use} command in the .do file above the replace statements.{p_end}
{pmore} {it:*See {help assertlist_replace##note:NOTE} for additional information regarding {cmd:DATASET1}.} {p_end} 
{marker dataset2}
{pstd} {bf:DATASET2} - Name for the new dataset that will contain the changes from the replace commands. If not populated, default value is {bf:dataset_with_replaced_values}. {p_end}
{pmore} {it:*See {help assertlist_replace##note:NOTE} for additional information regarding {cmd:DATASET2}.} {p_end}

{marker note}	   
{pstd} {bf:NOTE: The input for {it:EXCEL}, {it:DOFILE}, {it:DATASET1} and {it:DATASET2} can include just the file name (goes to current folder) or a folder} 
        {bf: path and file name. Do {it:NOT} include double quotes around the path and filename for output excel file. Do not include file extensions .dta or .do in these options. }{p_end}


{hline}

{title:Authors}
{p}

Mary Kay Trimner & Dale Rhoda, Biostat Global Consulting

Email {browse "mailto:Dale.Rhoda@biostatglobal.com":Dale.Rhoda@biostatglobal.com}

Biostat Global Consulting created two additional programs that go along with {cmd:assertlist_replace} : 
{pstd} {help assertlist} : Initial program that must be run prior to running {cmd:assertlist_replace}. {p_end}
{pmore} {cmd:Assertlist}  List observations that contradict an assert command. {p_end}

{pstd} {help assertlist_cleanup} : Optional program that can be run prior to running {cmd:assertlist_replace}. {p_end}
{pmore}{cmd:Assertlist_cleanup} Cleans up excel file generated by assertlist. {p_end}


{title:See Also}
{help assert}
{help assertlist}
{help assertlist_cleanup}
