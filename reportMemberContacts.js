function reportMemberContacts(stmt) {
	makeReport(stmt, {
		sheetName: "Member-Contacts",
		query: `
      SELECT 
        first_name AS "First Name",
        last_name AS "Last Name",
        nickname AS "Facebook Name",
        cc_member AS "Is Member",
        \`admin-phone-number\` AS "Phone Number",
        user_email AS "Email"
      FROM jtl_member_db
      WHERE \`admin-phone-number\` IS NOT NULL 
        AND \`admin-phone-number\` != ''
      ORDER BY first_name ASC, last_name ASC
    `,
		formatting: [
			{ type: "width", column: "First Name", width: 150 },
			{ type: "width", column: "Last Name", width: 150 },
			{ type: "width", column: "Phone Number", width: 120 },
			{ type: "width", column: "Email", width: 200 },
		],
	});
}
