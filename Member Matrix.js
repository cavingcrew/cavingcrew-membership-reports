
function readData() {
 var conn = Jdbc.getConnection(url, username, password);
 var stmt = conn.createStatement();
 

 var spreadsheet = SpreadsheetApp.getActive();
 var sheet = spreadsheet.getSheetByName('BCA-CIM-Proforma');
 //var cell = sheet.getRange('B4').getValues();

var results = stmt.executeQuery('select distinct first_name as "Forenames",`last_name` as "Surname", "" as "Previous Name",`admin-bca-number` as "Membership Number", (SELECT (CASE WHEN `admin-other-club-name`<>"" THEN `admin-other-club-name` ELSE "The Caving Crew" END)) as  "Primary Club Name", `membership_joining_date` "Joining Date", (SELECT (CASE WHEN `admin-other-club-name`<>"" THEN "£0" ELSE "£20" END)) as "Fee Paid", (SELECT (CASE WHEN `admin-other-club-name`<>"" THEN "AN" ELSE "C" END)) as "Insurance Status", `user_email` "Email", `admin-personal-pronouns` "Gender" , `admin-personal-year-of-birth` "Year Of Birth", `billing_address_1` "Address 1", `billing_address_2` "Address 2", " " , `billing_city` "Town","" as "County", billing_postcode "Postcode","UK", id from jtl_member_db where `cc_member`="yes" order by STR_TO_DATE(membership_joining_date, "%d/%m/%Y") ASC,`admin-bca-number`ASC, `first_name` ASC')


 //var results = stmt.executeQuery('select distinct  `first_name` "Forenames",`last_name` "Surname", `admin-personal-year-of-birth` "Year of Birth", `shipping_address_1` "Address 1",`shipping_address_2` "Address 2",`shipping_city` "Town", `shipping_postcode` "Postcode",`billing_email` "Email",`admin-phone-number` "Home Tel.",`admin-bca-number` "BCA Number",pd.user_id "Club Ref", pd.order_id "Order Ref" from jtl_member_db db JOIN jtl_order_product_customer_lookup pd on pd.user_id = db.id where status in ("wc-processing", "wc-completed") AND YEAR(pd.order_created)>YEAR(CURDATE())-3 order by `first_name` ASC')
  //console.log(results);
   sheet.clearContents();

appendToSheet(sheet,results);
getMembersToProcess();
getMembersProcessed();

}

function getMembersToProcess(){

 var conn = Jdbc.getConnection(url, username, password);
 var stmt = conn.createStatement();

// to_process Report
//
 var spreadsheet = SpreadsheetApp.getActive();
 var sheet = spreadsheet.getSheetByName('To Process');
 sheet.clearContents();
 sheet.appendRow(["Members to process"]);


// start of to_process function
function to_process(querystring, title)
{
sheet.appendRow([title]);
//sheet.getLastRow(row+1,1).setFontWeight("bold");
//sheet.getRange(row+1,1).setFontWeight("bold");
//var lastRow = sheet.getLastRow(row+1,1);
var results = stmt.executeQuery('select distinct first_name as "Forenames",`last_name` as "Surname", "" as "Previous Name",`admin-bca-number` as "Membership Number", (SELECT (CASE WHEN `admin-other-club-name`<>"" THEN `admin-other-club-name` ELSE "The Caving Crew" END)) as  "Primary Club Name", `membership_joining_date` "Joining Date", (SELECT (CASE WHEN `admin-other-club-name`<>"" THEN "£0" ELSE "£20" END)) as "Fee Paid", (SELECT (CASE WHEN `admin-other-club-name`<>"" THEN "AN" ELSE "C" END)) as "Insurance Status", `user_email` "Email", `admin-personal-pronouns` "Gender" , `admin-personal-year-of-birth` "Year Of Birth", `billing_address_1` "Address 1", `billing_address_2` "Address 2", " " , `billing_city` "Town","" as "County", billing_postcode "Postcode","UK", id, (SELECT CONCAT("https://www.cavingcrew.com/wp-admin/post.php?post=",pd.order_id,"&action=edit") ) as "Order Edit" from jtl_member_db where `cc_member`="yes" AND `admin-bca-number IS NULL` order by STR_TO_DATE(membership_joining_date, "%d/%m/%Y") ASC,`admin-bca-number`ASC, `first_name` ASC')

// var results = stmt.executeQuery('select distinct  `first_name` "Forenames",`last_name` "Surname", `shipping_address_1` "Address 1",`shipping_address_2` "Address 2",`shipping_city` "Town", `shipping_postcode` "Postcode",`billing_email` "Email",`admin-phone-number` "Home Tel.",pd.order_item_name "Membership Name", pd.variation_id "Variation ID", user_id "User ID", pd.order_id "Order ID",  from jtl_member_db db JOIN jtl_order_product_customer_lookup pd on pd.user_id = db.id where product_id=548 AND status="wc-processing" AND YEAR(pd.order_created)>YEAR(CURDATE())-1  AND  `admin-bca-number`="" order by `first_name` ASC');
  //console.log(results);
appendToSheet(sheet,results);




} //end of to_process

// full options
//help online beforehand,help at sign-in,help around announcements and cake time,do announcements,help online afterwards,be event director for the evening
 sheet.clearContents();

to_process("548", "Members to Process");



}

function getMembersProcessed(){

// Processed Report
//
 var spreadsheet = SpreadsheetApp.getActive();
 var sheet = spreadsheet.getSheetByName('Processed');
 sheet.clearContents();
 sheet.appendRow(["Processed Members"]);


// start of Processed function
function processed(querystring, title)
{
sheet.appendRow([title]);
//sheet.getLastRow(row+1,1).setFontWeight("bold");
//sheet.getRange(row+1,1).setFontWeight("bold");
//var lastRow = sheet.getLastRow(row+1,1);

 var results = stmt.executeQuery('select distinct  `first_name` "Forenames",`last_name` "Surname", pd.order_item_name "Membership Name", pd.variation_id "Variation ID", user_id "User ID", pd.order_id "Order ID" from jtl_member_db db JOIN jtl_order_product_customer_lookup pd on pd.user_id = db.id where product_id=548 AND status="wc-completed" AND YEAR(pd.order_created)>YEAR(CURDATE())-1 AND  `admin-bca-number`<>""  order by `first_name` ASC');
  //console.log(results);
appendToSheet(sheet,results);




} //end of processed

// full options
//help online beforehand,help at sign-in,help around announcements and cake time,do announcements,help online afterwards,be event director for the evening
 sheet.clearContents();

processed("2360", "Benefactor Members");
processed("2361", "Members with Mugs");
processed("2357", "Regular Members");
//volunteering("online afterwards", "Online Afterwards");
//volunteering("do announcements", "Do Announcements");
//volunteering("be event dir", "Be Event Director");

//var range = SpreadsheetApp.getActive().getRange("Volunteering!C2:C150");
//range.insertCheckboxes();
//sheet.autoResizeColumns(1, numCols+1);






//close SQL
results.close();
stmt.close();



//end read data function
} 


//ScriptApp.newTrigger('readData')
//.timeBased()
//.everyMinutes(30)
//.create();
