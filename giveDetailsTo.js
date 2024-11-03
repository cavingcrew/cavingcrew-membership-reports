function unMembership() {
	giveCompetency("cc_member", "no");
}

function giveMembership() {
	giveCompetency("cc_member", "yes");
}

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
				ui.alert(
					'Please use this function from either the "To Process" or "BCA-CIM-Proforma" sheets',
				);
		}
	} catch (error) {
		ui.alert("Error: " + error.message);
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

function giveCompetency(meta_key, meta_value) {
	var spreadsheet = SpreadsheetApp.getActive();
	var sheet = spreadsheet.getSheetByName("BCA-CIM-Proforma");
	var active_range = sheet.getActiveRange();
	var currentRow = active_range.getRowIndex();

	if (currentRow <= 1) {
		Browser.msgBox("Select an actual signup", Browser.Buttons.OK);
		return;
	}
	if (currentRow >= 100) {
		Browser.msgBox("Select an actual signup", Browser.Buttons.OK);
		return;
	}

	// Get headers from row 1
	var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

	// Find the index of the 'id' column (adding 1 since array index starts at 0 but sheets start at 1)
	var idColumnIndex = headers.indexOf("id") + 1;

	if (idColumnIndex === 0) {
		Browser.msgBox("Could not find id column", Browser.Buttons.OK);
		return;
	}

	var user_id = sheet.getRange(currentRow, idColumnIndex, 1, 1).getValue();
	var first_name = sheet.getRange(currentRow, 1, 1, 1).getValue();

	if (user_id === "" || user_id === "user_id") {
		Browser.msgBox("No User ID Found", Browser.Buttons.OK);
		return;
	}

	if (
		Browser.msgBox(
			"Update details for " + first_name + "? \n User " + user_id,
			Browser.Buttons.OK_CANCEL,
		) == "ok"
	) {
		const cc_attendance_setter = Session.getActiveUser().getEmail();

		//let metakey = "milestones_3_badge"
		//let metavalue = "given"
		const datetime = Date.now();
		//Logger.log(datetime);

		const meta_key_given_at = meta_key + "_marked_given_at";
		const meta_key_given_by = meta_key + "_marked_given_by";

		var data = {
			meta_data: [
				{ key: meta_key, value: meta_value },
				{ key: meta_key_given_at, value: datetime },
				{ key: meta_key_given_by, value: cc_attendance_setter },
			],
		};

		let returndata = pokeToWooUserMeta(data, user_id); //returns JSON object
		returndata = returndata.getContentText();
		returndata = JSON.parse(returndata);

		//Logger.log("type " + typeof(returndata)); // type object
		//Logger.log(returndata.data); // Logging output too large. Truncating output. {"id":52,"date_created":"2021-08-27T23:24:39","date_created_gmt":"2021-08-27T22:24:39","date_modified":"2022-11-23T11:51:41", etc etc etc
		//Logger.log(returndata[0]); //null
		//Logger.log("type " + typeof(returndata.id)); //type undefined
		//Logger.log(returndata.data.id); //null

		const search = returndata.meta_data.find(
			({ key }) => key == meta_key,
		)?.value;

		//Logger.log(search);

		if (search === meta_value) {
			sheet.getRange(currentRow, 2, 1, 1).setValue("Given"); // paste the blank variables into the cells to delete contents
			sheet.getRange(currentRow, 15, 1, 1).setValue(meta_key);
			sheet.getRange(currentRow, 16, 1, 1).setValue(meta_value); // paste the blank variables into the cells to delete contents

			return meta_value;
		} else {
			Logger.log("ERROR" + search);

			return "ERROR" + search;
		}
		return "ERROR";
	}

	//Logger.log(returnvalue.meta_data);
	//Logger.log(" type " + typeof(returnvalue.meta_data));

	//Logger.log(JSON.parse(returnvalue));
	//Logger.log(JSON.stringify(returnvalue.meta_data));

	//return returndata;
}
