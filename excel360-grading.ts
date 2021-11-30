// Script: NewEvalSheet
function main(workbook: ExcelScript.Workbook) {
	let namecell = "E2"
	let selectedSheet = workbook.getWorksheet("template");
	// Duplicate worksheet
	let mysheet = selectedSheet.copy(ExcelScript.WorksheetPositionType.end);
	// Rename worksheet to whatever is in the namecell
	let student = mysheet.getRange(namecell).getText();
	console.log("Creating student sheet "+student)
	mysheet.setName(student);
}
