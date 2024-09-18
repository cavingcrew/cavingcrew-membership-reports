

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