function reportCancellation(stmt) {
    makeReport(stmt, {
        sheetName: "Cancelled Members",
        query: `                                                                                                                                                                                                                      
       SELECT DISTINCT                                                                                                                                                                                                             
         first_name AS "Forenames",                                                                                                                                                                                                
         last_name AS "Surname",                                                                                                                                                                                                   
         "" AS "Previous Name",                                                                                                                                                                                                    
         \`admin-bca-number\` AS "Membership Number",                                                                                                                                                                                  
         (SELECT (CASE WHEN \`admin-other-club-name\`<>"" THEN \`admin-other-club-name\` ELSE "The Caving Crew" END)) AS "Primary Club Name",                                                                                              
         membership_cancellation_date AS "Cancellation Date",                                                                                                                                                                      
         cc_member AS "Crew Member?",                                                                                                                                                                                              
         (SELECT (CASE WHEN \`admin-other-club-name\`<>"" THEN "AN" ELSE "C" END)) AS "Insurance Status",                                                                                                                              
         user_email AS "Email",                                                                                                                                                                                                    
         \`admin-personal-pronouns\` AS "Gender",                                                                                                                                                                                      
         \`admin-personal-year-of-birth\` AS "Year Of Birth",                                                                                                                                                                          
         billing_address_1 AS "Address 1",                                                                                                                                                                                         
         billing_address_2 AS "Address 2",                                                                                                                                                                                         
         "" AS "Address 3",                                                                                                                                                                                                        
         billing_city AS "Town",                                                                                                                                                                                                   
         "" AS "County",                                                                                                                                                                                                           
         billing_postcode AS "Postcode",                                                                                                                                                                                           
         "UK",                                                                                                                                                                                                                     
         id                                                                                                                                                                                                                        
       FROM jtl_member_db                                                                                                                                                                                                          
       WHERE (cc_member="no" OR cc_member="expired")                                                                                                                                                                               
       ORDER BY STR_TO_DATE(membership_cancellation_date, '%d/%m/%Y') DESC                                                                                                                                                                                  
     `,
        formatting: [
            { type: 'wrap', column: "Address 1" },
            { type: 'wrap', column: "Address 2" },
            { type: 'wrap', column: "Town" },
            { type: 'wrap', column: "Postcode" },
            { type: 'color', column: "Crew Member?", search: "no", color: colors.lightRed },
            { type: 'color', column: "Crew Member?", search: "expired", color: colors.lightYellow },
        ]
    });
}
