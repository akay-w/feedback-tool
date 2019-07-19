# feedback-tool
Translation feedback tool for SDL Trados

Created in Python 3.7.1

Non-standard dependent libraries: xlsxwriter

This tool can be used to compare original .sdlxliff files delivered by an outsourcer with the final edited .sdlxliff files. It generates an Excel file that shows all of the changes in red, with space for comments. The idea was to make it a bit easier to give feedback to translators. It currently works with Trados 2015 and above.

To use the tool, run the feedback_tool_2_gui.py file. After launching the tool, for a given project, use the buttons to select a folder containing the original sdlxliff files, a folder containing the final sdlxliff files, a save location, and the type of target language (European/Asian). Note that only folders can be selected. The only difference between the European/Asian options is that for the European language the tool will make sure to put spaces between tokens when showing the changes. Optionally, enter a project no./name, which will be included in the created Excel report name. Click Run to run the tool, and Exit to close it.

The tool will compare all of the original and final sdxliff files, and output an Excel report showing the differences. Segments in which there are no changes will be hidden by default. If there are changes in the source (due to source edits being used), the final source will be displayed in a additional column, otherwise this column is hidden by default. Changes will be shown in red (deletions from original file and additions to final file).

Known issues: 

If the entire target segment has been changed, it doesn't correctly print the changed target segment, because the function to write rich text needs to have two arguments. I need to fix that.

All tags are currently ignored. I'd like to add those in later, so that changes in formatting/placeholder tags can be seen.

Doesn't work correctly if there are un-translated segments. I need to look into that further.
