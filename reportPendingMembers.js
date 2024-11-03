function reportPendingMembers(stmt) {
  makeReport(stmt, {
    sheetName: "Pending Members",
    query: `
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
    `,
    formatting: [
      { type: 'wrap', column: "Address1" },
      { type: 'wrap', column: "Address2" },
      { type: 'wrap', column: "Town" },
      { type: 'wrap', column: "PostCode" },
      { type: 'numberFormat', column: "DOB*", format: "dd/mm/yyyy" }
    ]
  });
}
