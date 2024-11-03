function reportJustGoMembers(stmt) {
	makeReport(stmt, {
		sheetName: "JustGo Members",
		query: `
      SELECT DISTINCT
        \`admin-bca-number\` AS "MID",
        first_name AS "Firstname",
        last_name AS "Surname",
        CONCAT('01/01/', \`admin-personal-year-of-birth\`) AS "DOB",
        CASE 
          WHEN \`admin-personal-pronouns\` = 'm' THEN 'Male'
          WHEN \`admin-personal-pronouns\` = 'f' THEN 'Female'
          ELSE ''
        END AS "Gender",
        \`admin-phone-number\` AS "Contact Number",
        SUBSTRING_INDEX(\`admin-emergency-contact-name\`, ' ', 1) AS "Emergency Firstname",
        SUBSTRING_INDEX(\`admin-emergency-contact-name\`, ' ', -1) AS "Emergency Surname",
        '' AS "Emergency Relation",
        \`admin-emergency-contact-phone\` AS "Emergency Number",
        '' AS "Emergency Email",
        user_email AS "Email Address",
        billing_address_1 AS "Address Line 1",
        billing_address_2 AS "Address Line 2",
        billing_city AS "Town",
        '' AS "County",
        billing_postcode AS "Postcode",
        COALESCE(billing_country, 'United Kingdom') AS "Country",
        CASE 
          WHEN \`admin-other-club-name\` = '' THEN 'True'
          ELSE 'False'
        END AS "Is Primary Club",
        CASE 
          WHEN \`admin-other-club-name\` = '' THEN 'The Caving Crew'
          ELSE ''
        END AS "Primary Club",
        CASE 
          WHEN \`admin-other-club-name\` != '' THEN 'The Caving Crew'
          ELSE ''
        END AS "Additional Clubs",
        'Caving Member' AS "Membership Type",
        DATE_FORMAT(
          DATE_ADD(STR_TO_DATE(membership_joining_date, '%d/%m/%Y'), INTERVAL 1 YEAR),
          '%d/%m/%Y'
        ) AS "Membership Creation Date",
        '31/12/2024' AS "Expiry Date",
        CASE 
          WHEN cc_member = 'yes' THEN 'Registered'
          ELSE 'Unregistered'
        END AS "Club Member Status"
      FROM jtl_member_db
      WHERE cc_member = 'yes'
      AND \`admin-bca-number\` IS NOT NULL 
      AND \`admin-bca-number\` != ''
      ORDER BY \`admin-bca-number\` ASC
    `,
		formatting: [
			{ type: "wrap", column: "Address Line 1" },
			{ type: "wrap", column: "Address Line 2" },
			{ type: "wrap", column: "Town" },
			{ type: "wrap", column: "Postcode" },
			{ type: "numberFormat", column: "DOB", format: "dd/mm/yyyy" },
			{
				type: "numberFormat",
				column: "Membership Creation Date",
				format: "dd/mm/yyyy",
			},
			{ type: "numberFormat", column: "Expiry Date", format: "dd/mm/yyyy" },
		],
	});
}
