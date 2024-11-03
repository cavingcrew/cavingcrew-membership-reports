function updateExternalPendingMembers(stmt) {
	// Define the external spreadsheet and sheet
	const externalSpreadsheetId = "1DUEckpPQidMZDWFZ0OnJxCroLiLjh2gPgrLk";
	const externalSheet = SpreadsheetApp.openById(
		externalSpreadsheetId,
	).getSheetByName("members");

	// Clear rows 2-10
	externalSheet.getRange(2, 1, 9, externalSheet.getLastColumn()).clearContent();

	// Run the same query as reportPendingMembers
	const results = stmt.executeQuery(`
    SELECT DISTINCT
      first_name AS "Firstname*",
      last_name AS "Lastname*",
      user_email AS "EmailAddress*",
      CONCAT('01/01/', \`admin-personal-year-of-birth\`) AS "DOB*",
      user_login AS "Username*",
      \`admin-personal-pronouns\` AS "Gender",
      '' AS "Title",
      billing_address_1 AS "Address1",
      billing_address_2 AS "Address2", 
      billing_city AS "Town",
      '' AS "County",
      billing_postcode AS "PostCode",
      billing_country AS "Country",
      \`admin-phone-number\` AS "Mobile Telephone",
      billing_phone AS "Home Telephone",
      SUBSTRING_INDEX(\`admin-emergency-contact-name\`, ' ', 1) AS "Emergency Contact First Name",
      SUBSTRING_INDEX(\`admin-emergency-contact-name\`, ' ', -1) AS "Emergency Contact Surname",
      '' AS "Emergency Contact Relationship",
      \`admin-emergency-contact-phone\` AS "Emergency Contact Number",
      '' AS "Emergency Contact Email Address",
      '' AS "Parent FirstName",
      '' AS "Parent Surname",
      '' AS "Parent EmailAddress"
    FROM jtl_member_db
    WHERE cc_member='yes' 
    AND (\`admin-bca-number\` IS NULL OR \`admin-bca-number\` = '')
    ORDER BY first_name ASC, last_name ASC
  `);

	// Convert results to array, skipping the header row
	const rows = [];
	const metaData = results.getMetaData();
	const numCols = metaData.getColumnCount();

	while (results.next()) {
		const row = [];
		for (let col = 0; col < numCols; col++) {
			row.push(results.getString(col + 1));
		}
		rows.push(row);
	}

	// If we have results, write them starting at row 2
	if (rows.length > 0) {
		externalSheet.getRange(2, 1, rows.length, numCols).setValues(rows);
	}

	results.close();
}
