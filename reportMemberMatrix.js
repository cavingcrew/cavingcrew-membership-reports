function reportMemberMatrix() {
 const conn = Jdbc.getConnection(url, username, password);
 const stmt = conn.createStatement();

 makeReport(stmt, {
  sheetName: "BCA-CIM-Proforma",
  query: `                                                                                                                                                                                                                      
       SELECT DISTINCT                                                                                                                                                                                                             
         first_name AS "Forenames",                                                                                                                                                                                                
         last_name AS "Surname",                                                                                                                                                                                                   
         "" AS "Previous Name",                                                                                                                                                                                                    
         admin-bca-number AS "Membership Number",                                                                                                                                                                                  
         (SELECT (CASE WHEN admin-other-club-name<>"" THEN admin-other-club-name ELSE "The Caving Crew" END)) AS "Primary Club Name",                                                                                              
         membership_joining_date AS "Joining Date",                                                                                                                                                                                
         (SELECT (CASE WHEN admin-other-club-name<>"" THEN "£0" ELSE "£20" END)) AS "Fee Paid",                                                                                                                                    
         (SELECT (CASE WHEN admin-other-club-name<>"" THEN "AN" ELSE "C" END)) AS "Insurance Status",                                                                                                                              
         user_email AS "Email",                                                                                                                                                                                                    
         admin-personal-pronouns AS "Gender",                                                                                                                                                                                      
         admin-personal-year-of-birth AS "Year Of Birth",                                                                                                                                                                          
         billing_address_1 AS "Address 1",                                                                                                                                                                                         
         billing_address_2 AS "Address 2",                                                                                                                                                                                         
         " ",                                                                                                                                                                                                                      
         billing_city AS "Town",                                                                                                                                                                                                   
         "" AS "County",                                                                                                                                                                                                           
         billing_postcode AS "Postcode",                                                                                                                                                                                           
         "UK",                                                                                                                                                                                                                     
         id                                                                                                                                                                                                                        
       FROM jtl_member_db                                                                                                                                                                                                          
       WHERE cc_member="yes"                                                                                                                                                                                                       
       ORDER BY STR_TO_DATE(membership_joining_date, "%d/%m/%Y") ASC, admin-bca-number ASC, first_name ASC                                                                                                                         
     `,
  formatting: [
   { type: 'wrap', column: "Address 1" },
   { type: 'wrap', column: "Address 2" },
   { type: 'wrap', column: "Town" },
   { type: 'wrap', column: "Postcode" },
  ]
 });

 getMembersToProcess(stmt);
 getMembersProcessed(stmt);

 stmt.close();
 conn.close();
}

function getMembersToProcess(stmt) {
 makeReport(stmt, {
  sheetName: "To Process",
  query: `                                                                                                                                                                                                                      
       SELECT DISTINCT                                                                                                                                                                                                             
         first_name AS "Forenames",                                                                                                                                                                                                
         last_name AS "Surname",                                                                                                                                                                                                   
         "" AS "Previous Name",                                                                                                                                                                                                    
         admin-bca-number AS "Membership Number",                                                                                                                                                                                  
         (SELECT (CASE WHEN admin-other-club-name<>"" THEN admin-other-club-name ELSE "The Caving Crew" END)) AS "Primary Club Name",                                                                                              
         membership_joining_date AS "Joining Date",                                                                                                                                                                                
         (SELECT (CASE WHEN admin-other-club-name<>"" THEN "£0" ELSE "£20" END)) AS "Fee Paid",                                                                                                                                    
         (SELECT (CASE WHEN admin-other-club-name<>"" THEN "AN" ELSE "C" END)) AS "Insurance Status",                                                                                                                              
         user_email AS "Email",                                                                                                                                                                                                    
         admin-personal-pronouns AS "Gender",                                                                                                                                                                                      
         admin-personal-year-of-birth AS "Year Of Birth",                                                                                                                                                                          
         billing_address_1 AS "Address 1",                                                                                                                                                                                         
         billing_address_2 AS "Address 2",                                                                                                                                                                                         
         " ",                                                                                                                                                                                                                      
         billing_city AS "Town",                                                                                                                                                                                                   
         "" AS "County",                                                                                                                                                                                                           
         billing_postcode AS "Postcode",                                                                                                                                                                                           
         "UK",                                                                                                                                                                                                                     
         id,                                                                                                                                                                                                                       
         (SELECT CONCAT("https://www.cavingcrew.com/wp-admin/post.php?post=",pd.order_id,"&action=edit")) AS "Order Edit"                                                                                                          
       FROM jtl_member_db                                                                                                                                                                                                          
       LEFT JOIN jtl_order_product_customer_lookup pd ON pd.user_id = jtl_member_db.id                                                                                                                                             
       WHERE cc_member="yes" AND admin-bca-number IS NULL                                                                                                                                                                          
       ORDER BY STR_TO_DATE(membership_joining_date, "%d/%m/%Y") ASC, admin-bca-number ASC, first_name ASC                                                                                                                         
     `,
  formatting: [
   { type: 'wrap', column: "Address 1" },
   { type: 'wrap', column: "Address 2" },
   { type: 'wrap', column: "Town" },
   { type: 'wrap', column: "Postcode" },
  ]
 });
}

function getMembersProcessed(stmt) {
 makeReport(stmt, {
  sheetName: "Processed",
  query: `                                                                                                                                                                                                                      
       SELECT DISTINCT                                                                                                                                                                                                             
         first_name AS "Forenames",                                                                                                                                                                                                
         last_name AS "Surname",                                                                                                                                                                                                   
         pd.order_item_name AS "Membership Name",                                                                                                                                                                                  
         pd.variation_id AS "Variation ID",                                                                                                                                                                                        
         user_id AS "User ID",                                                                                                                                                                                                     
         pd.order_id AS "Order ID"                                                                                                                                                                                                 
       FROM jtl_member_db db                                                                                                                                                                                                       
       JOIN jtl_order_product_customer_lookup pd ON pd.user_id = db.id                                                                                                                                                             
       WHERE product_id=548 AND status="wc-completed" AND YEAR(pd.order_created)>YEAR(CURDATE())-1 AND admin-bca-number<>""                                                                                                        
       ORDER BY first_name ASC                                                                                                                                                                                                     
     `
 });
}