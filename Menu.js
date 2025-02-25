function onOpen() {
	const ui = SpreadsheetApp.getUi();
	ui.createMenu("Mark Membership")
		.addItem("Remove Membership", "unMembership")
		.addItem("Give Membership", "giveMembership")
		.addToUi();

	ui.createMenu("Input Data")
		.addItem("Input BCA Number", "setBCANumber")
		.addItem("Input BirthYear", "setBirthYear")
		.addToUi();

	ui.createMenu("Member Management")
		.addItem("Edit Member Details", "editMemberDetails")
		.addToUi();

	ui.createMenu("Refresh Matrix")
		.addItem("Refresh All", "readData")
		.addSeparator()
		.addItem("Refresh BCA-CIM-Proforma", "refreshBCACIMProforma")
		.addItem("Refresh Members To Process", "refreshMembersToProcess")
		.addItem("Refresh Processed Members", "refreshMembersProcessed")
		.addSeparator()
		.addItem("Refresh Cancellation", "refreshCancellation")
		.addItem("Refresh Pending Members", "refreshPendingMembers")
		.addItem("Refresh JustGo Members", "refreshJustGoMembers")
		.addItem(
			"Refresh External Pending Members",
			"refreshExternalPendingMembers",
		)
		.addToUi();
}

function refreshBCACIMProforma() {
	const conn = Jdbc.getConnection(url, username, password);
	const stmt = conn.createStatement();
	reportBCACIMProforma(stmt);
	stmt.close();
	conn.close();
}

function refreshMembersToProcess() {
	const conn = Jdbc.getConnection(url, username, password);
	const stmt = conn.createStatement();
	reportMembersToProcess(stmt);
	stmt.close();
	conn.close();
}

function refreshMembersProcessed() {
	const conn = Jdbc.getConnection(url, username, password);
	const stmt = conn.createStatement();
	reportMembersProcessed(stmt);
	stmt.close();
	conn.close();
}

function refreshCancellation() {
	const conn = Jdbc.getConnection(url, username, password);
	const stmt = conn.createStatement();
	reportCancellation(stmt);
	stmt.close();
	conn.close();
}

function refreshPendingMembers() {
	const conn = Jdbc.getConnection(url, username, password);
	const stmt = conn.createStatement();
	reportPendingMembers(stmt);
	stmt.close();
	conn.close();
}

function refreshJustGoMembers() {
	const conn = Jdbc.getConnection(url, username, password);
	const stmt = conn.createStatement();
	reportJustGoMembers(stmt);
	stmt.close();
	conn.close();
}

function refreshExternalPendingMembers() {
	const conn = Jdbc.getConnection(url, username, password);
	const stmt = conn.createStatement();
	updateExternalPendingMembers(stmt);
	stmt.close();
	conn.close();
}
