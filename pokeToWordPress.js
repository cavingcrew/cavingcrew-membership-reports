function updateBCAMembershipDetails(userId, orderId, bcaNumber) {
  const { user, datetime } = getCurrentUserAndTime();

  const userMetaData = {
    "meta_data": [
      {
        "key": "admin-bca-number",
        "value": bcaNumber
      },
      {
        "key": "admin-bca-number_marked_given_at",
        "value": datetime
      },
      {
        "key": "admin-bca-number_marked_given_by",
        "value": user
      }
    ]
  };

  try {
    const userResponse = pokeToWooUserMeta(userMetaData, userId);
    if (!userResponse.getResponseCode() === 200) {
      throw new Error('Failed to update user metadata');
    }

    const orderData = {
      "status": "completed",
      "meta_data": [
        {
          "key": "admin_order_processed_at",
          "value": datetime
        },
        {
          "key": "admin_order_processed_by",
          "value": user
        }
      ]
    };

    const orderResponse = pokeToWordPressOrders(orderData, orderId);
    if (!orderResponse.getResponseCode() === 200) {
      throw new Error('Failed to update order status');
    }

    return true;
  } catch (error) {
    throw error;
  }
}

function pokeToWordPressOrders(data, user_id) {

  //console.log("Wordpress " + data);

  let encodedAuthInformation = Utilities.base64Encode(apiusername + ":" + apipassword);
  let headers = { "Authorization": "Basic " + encodedAuthInformation };
  let options = {
    'method': 'post',
    'contentType': 'application/json',
    'headers': headers,  // Convert the JavaScript object to a JSON string.
    'payload': JSON.stringify(data)
  };
  apiurl = "https://www." + apidomain + "/wp-json/wc/v3/orders/" + user_id

  return response = UrlFetchApp.fetch(apiurl, options);
 // console.log(response);
}

function pokeToWordPressProducts(data, product_id) {

  //console.log("Wordpress " + data);

  let encodedAuthInformation = Utilities.base64Encode(apiusername + ":" + apipassword);
  let headers = { "Authorization": "Basic " + encodedAuthInformation };
  let options = {
    'method': 'post',
    'contentType': 'application/json',
    'headers': headers,  // Convert the JavaScript object to a JSON string.
    'payload': JSON.stringify(data)
  };
  apiurl = "https://www." + apidomain + "/wp-json/wc/v3/products/" + product_id

  return response = UrlFetchApp.fetch(apiurl, options);
 // console.log(response);
}

function pokeToWooUserMeta(data, user_id) {

  //console.log("Wordpress " + data);

  let encodedAuthInformation = Utilities.base64Encode(apiusername + ":" + apipassword);
  let headers = { "Authorization": "Basic " + encodedAuthInformation };
  let options = {
    'method': 'post',
    'contentType': 'application/json',
    'headers': headers,  // Convert the JavaScript object to a JSON string.
    'payload': JSON.stringify(data)
  };
  apiurl = "https://www." + apidomain + "/wp-json/wc/v3/customers/" + user_id

  return response = UrlFetchApp.fetch(apiurl, options);
 // console.log(response);
}
