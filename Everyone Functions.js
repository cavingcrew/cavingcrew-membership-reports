//      .addItem('Finalise Assignments', 'test')
// MarkAttended and Close
// https://woocommerce.github.io/woocommerce-rest-api-docs/#batch-update-orders

// Get all IDs
// SELECT order_id  FROM wp_order_product_customer_lookup where product_id=$id AND status='wc-processing' AND cc_attendance='pending'
// Make on hold

function sendAllCragAssignments(tellme) {

  const action = tellme
  var sconn = Jdbc.getConnection(url, username, password);
  var sstmt = sconn.createStatement();

  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getSheetByName('Dashboard');
  var product_id = sheet.getRange('B5').getValues();
  //var product_id = "3103";

  var order_results = sstmt.executeQuery('SELECT distinct order_id from wp_order_product_customer_lookup where product_id="' + product_id + '"  AND `cc_location`="' + cc_location + '" AND status="wc-processing" AND cc_attendance="pending" LIMIT 99');
  let active_user = Session.getActiveUser().getEmail();
  let currentUnixTime = Date.now();
  //var order_results = sstmt.executeQuery('select "3575"');

  while (order_results.next()) {

    scores_arr = [];
    for (let col = 0; col < 1; col++) {
      scores_arr.push(order_results.getString(col + 1));
    }
    console.log(scores_arr);


    let order_id = scores_arr[0];


    if (action === "close") {

            var data = {
              "status": "completed",
              "meta_data": [
                {
                  "key": "cc_attendance_set_by",
                  "value": active_user
                },
                {
                  "key": "cc_attendance_set_at",
                  "value": currentUnixTime
                },
                {
                  "key": "cc_attendance",
                  "value": "attended"
                }

              ]
            }
            Logger.log(data);


            pokeToWordPressOrders(data, order_id);


 



    } /*else if (action === "location") {
      if (Browser.msgBox("This will finalise the locations and email everyone their locations immediately", Browser.Buttons.OK_CANCEL) == "ok") {
        if (Browser.msgBox("This should be done after 12:00 on Wed", Browser.Buttons.OK_CANCEL) == "ok") {
          if (Browser.msgBox("This cannot be undone", Browser.Buttons.OK_CANCEL) == "ok") {



            var data = {
              "id": order_id,
              "status": "on-hold",
              "meta_data": [
                {
                  "key": "cc_location_signed_off_by",
                  "value": active_user
                },
                {
                  "key": "cc_location_finalised_at",
                  "value": currentUnixTime
                }
              ]
            }

            Logger.log(data);


            pokeToWordPressOrders(data, order_id);

          }
        }
      }
    } else {
      console.log("All done")
    }
    //conn.close();

    //console.log("done")
  }


  if (action === "close") {


    var data = {
      "status": "private",
      "meta_data": [
        {
          "key": "cc_post_set_private_set_by",
          "value": active_user
        },
        {
          "key": "cc_post_set_private_set_at",
          "value": currentUnixTime
        }

      ]

    }
    console.log("data");

    pokeToWordPressProducts(data, product_id);
  } */
else {
    console.log("All done")
  }



}}


//console.log(cell);


function markAttendedAndCloseEvent() {

        if (Browser.msgBox("This will mark all those who haven't been cancelled as ATTENDED and will close the event", Browser.Buttons.OK_CANCEL) == "ok") {
        if (Browser.msgBox("This should be done on Wednesday evening after climbing", Browser.Buttons.OK_CANCEL) == "ok") {

          if (Browser.msgBox("This cannot be undone", Browser.Buttons.OK_CANCEL) == "ok") {


  sendAllCragAssignments("close");
  //markEventComplete;
  /// Close the event 
          }}}
}
/*
function finaliseCragLocations() {
  sendAllCragAssignments("location");
}

*/

