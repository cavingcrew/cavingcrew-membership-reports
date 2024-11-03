function readData() {
	const conn = Jdbc.getConnection(url, username, password);
	const stmt = conn.createStatement();

	const reports = [
		reportBCACIMProforma,
		reportMembersToProcess,
		reportMembersProcessed,
		reportCancellation,
		reportMemberContacts,
		reportPendingMembers,
	];

	for (const report of reports) {
		report(stmt);
	}

	stmt.close();
	conn.close();
}
