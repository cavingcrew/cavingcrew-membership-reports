function setBCANumber() {
	const ui = SpreadsheetApp.getUi();
	const activeSheet = SpreadsheetApp.getActiveSheet();
	const currentRow = activeSheet.getActiveRange().getRowIndex();

	if (currentRow <= 1) {
		ui.alert("Please select a member row");
		return;
	}

	const input = ui.prompt(
		"Enter BCA Membership Number:",
		ui.ButtonSet.OK_CANCEL,
	);
	if (input.getSelectedButton() !== ui.Button.OK) return;

	const bcaNumber = input.getResponseText().trim();
	if (!bcaNumber) {
		ui.alert("Please enter a valid BCA number");
		return;
	}

	try {
		switch (activeSheet.getName()) {
			case "To Process":
				processBCANumberFromToProcess(activeSheet, currentRow, bcaNumber);
				break;
			case "BCA-CIM-Proforma":
				processBCANumberFromProforma(activeSheet, currentRow, bcaNumber);
				break;
			default:
				ui.alert('Please use this function from the "To Process" sheets');
		}
	} catch (error) {
		ui.alert(`Error: ${error.message}`);
	}
}

function processBCANumberFromToProcess(sheet, row, bcaNumber) {
	const idCol = findColumnIndex(sheet, "id");
	const orderEditCol = findColumnIndex(sheet, "Order Edit");
	const membershipNumberCol = findColumnIndex(sheet, "Membership Number");

	const userId = sheet.getRange(row, idCol).getValue();
	const orderEditUrl = sheet.getRange(row, orderEditCol).getValue();
	const orderId = extractOrderId(orderEditUrl);

	if (!userId || !orderId) {
		throw new Error("Could not find user ID or order ID");
	}

	updateBCAMembershipDetails(userId, orderId, bcaNumber);
	sheet.getRange(row, membershipNumberCol).setValue(bcaNumber);
}

function processBCANumberFromProforma(sheet, row, bcaNumber) {
	const idCol = findColumnIndex(sheet, "id");
	const membershipNumberCol = findColumnIndex(sheet, "Membership Number");
	const userId = sheet.getRange(row, idCol).getValue();

	if (!userId) {
		throw new Error("Could not find user ID");
	}

	const toProcessRow = findUserInSheet(userId, "To Process");
	if (!toProcessRow) {
		throw new Error("User not found in To Process sheet");
	}

	const toProcessSheet =
		SpreadsheetApp.getActive().getSheetByName("To Process");
	processBCANumberFromToProcess(toProcessSheet, toProcessRow, bcaNumber);

	sheet.getRange(row, membershipNumberCol).setValue(bcaNumber);
}
