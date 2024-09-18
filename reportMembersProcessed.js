function reportMembersProcessed() {
  const conn = Jdbc.getConnection(url, username, password);
  const stmt = conn.createStatement();

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

  stmt.close();
  conn.close();
}
