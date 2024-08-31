# My Personal Scripts
> *This repository contains my personal scripts that I use for my daily tasks. The scripts are written in various languages such as Python, JavaScript, VBScript, etc. The scripts are used for automating tasks, modifying web pages, and other purposes.*

## Description
* `ncbi_pmc_newUI_redirect.js` - redirect to new PubMed Central (PMC) experimental page.
* `ListWrongAnswers.bas` - custom function to list wrong answers in Excel. Explanation:
  * ListWrongAnswer(`start_calc_cell`, `start_ref_cell`, `wrong_questions`, `number_of_questions`, "")
  * `start_calc_cell` - The starting position of the calculation range.
  * `start_ref_cell` - The starting position of the reference range.
  * `wrong_questions` - Wrong answer questions.
  * `number_of_questions` - The number of questions.
* `ScoringUnweighted.bas` - custom function to calculate unweighted scoring in Excel (all questions have the same score). Explanation:
  * MatchScoreUnweighted(`start_calc_cell`, `start_ref_cell`, `number_of_questions`) (total value = 1)
  * `start_calc_cell` - The starting position of the calculation range.
  * `start_ref_cell` - The starting position of the reference range.
  * `number_of_questions` - The number of questions.
* `ScoringWeighted.bas` - custom function to calculate weighted scoring in Excel (this script is aligned with the Ministry of Education and Training (Vietnam)'s new scoring guidelines for the national high school exams)
  * MatchScoreWeighted(`start_calc_cell`, `start_ref_cell`, `number_of_questions`) (total value = 1)
  * `start_calc_cell` - The starting position of the calculation range.
  * `start_ref_cell` - The starting position of the reference range.
  * `number_of_questions` - The number of questions.
* `UpdateConstantCellRange.bas` - custom function to create array with constant cell range in Excel. Explanation:
  * UpdateConstantCellRange(`start_cell`, `number_of_cells_per_row`)
  * `start_cell` - The starting position of range.
  * `number_of_cells_per_row` - The number of cells per row.
  * Example: `=UpdateConstantCellRange(E$3,3)` will return an array containing the values of cells `E3`, `F3`, and `G3` (`E$3:G$3`).
* `UpdateDynamicCellRange.bas` - custom function to create array with dynamic cell range in Excel. Explanation:
  * UpdateDynamicCellRange(`start_cell`, `number_of_cells_per_row`)
  * `start_cell` - The starting position of range.
  * `number_of_cells_per_row` - The number of cells per row.
  * Example: `=UpdateDynamicCellRange(E3,3)` will return an array containing the values of cells `E3`, `F3`, `G3` (`E3:G3`).
## How to import Tampermonkey userscript into addon?
1) Open Tampermonkey dashboard by clicking on the Tampermonkey icon in the browser. Choose `Dashboard` option.
2) Go to `Utilities` tab.
3) In the `URL` field, paste the URL of the userscript.
4) Click on `Install` button.
5) Refresh the page to see the changes.

## How to import VBScript to Excel?
1) In Excel, press `Alt + F11` to open the VBA editor. Otherwise, you can go to `Developer` tab and click on `Visual Basic`.
2) In the VBA editor, go to `Insert` -> `Module` to create a new module.
3) Copy and paste the VBScript code into the module. 
> Otherwise, you can import the VBScript file by going to `File` → `Import File` and select the VBScript file.
4) Save VBA and then close the VBA editor and go back to Excel.
5) When save your Excel file, please save it as `.xlsm` file format.
