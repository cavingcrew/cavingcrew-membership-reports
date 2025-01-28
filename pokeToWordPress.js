function updateBCAMembershipDetails(userId, orderId, bcaNumber) {
	const { user, datetime } = getCurrentUserAndTime();

	const userMetaData = {
		meta_data: [
			{
				key: "admin-bca-number",
				value: bcaNumber,
			},
			{
				key: "admin-bca-number_marked_given_at",
				value: datetime,
			},
			{
				key: "admin-bca-number_marked_given_by",
				value: user,
			},
		],
	};

	const userResponse = pokeToWooUserMeta(userMetaData, userId);
	if (!userResponse.getResponseCode() === 200) {
		throw new Error("Failed to update user metadata");
	}

	const orderData = {
		status: "completed",
		meta_data: [
			{
				key: "admin_order_processed_at",
				value: datetime,
			},
			{
				key: "admin_order_processed_by",
				value: user,
			},
		],
	};

	const orderResponse = pokeToWordPressOrders(orderData, orderId);
	if (!orderResponse.getResponseCode() === 200) {
		throw new Error("Failed to update order status");
	}

	return true;
}

function pokeToWordPressOrders(data, user_id) {
	//console.log("Wordpress " + data);

	const encodedAuthInformation = Utilities.base64Encode(
		`${apiusername}:${apipassword}`,
	);
	const headers = { Authorization: `Basic ${encodedAuthInformation}` };
	const options = {
		method: "post",
		contentType: "application/json",
		headers: headers, // Convert the JavaScript object to a JSON string.
		payload: JSON.stringify(data),
	};
	const apiurl = `https://www.${apidomain}/wp-json/wc/v3/orders/${user_id}`;
	const response = UrlFetchApp.fetch(apiurl, options);
	return response;
	// console.log(response);
}

function pokeToWordPressProducts(data, product_id) {
	//console.log("Wordpress " + data);

	const encodedAuthInformation = Utilities.base64Encode(
		`${apiusername}:${apipassword}`,
	);
	const headers = { Authorization: `Basic ${encodedAuthInformation}` };
	const options = {
		method: "post",
		contentType: "application/json",
		headers: headers, // Convert the JavaScript object to a JSON string.
		payload: JSON.stringify(data),
	};
	const apiurl = `https://www.${apidomain}/wp-json/wc/v3/products/${product_id}`;
	const response = UrlFetchApp.fetch(apiurl, options);
	return response;
	// console.log(response);
}

function updateWooUser(userId, userData) {
	const { user, datetime } = getCurrentUserAndTime();

	// Transform the input data into the correct format
	const userMetaData = {
		meta_data: [],
	};

	// Add standard user fields if provided
	if (userData.first_name) userMetaData.first_name = userData.first_name;
	if (userData.last_name) userMetaData.last_name = userData.last_name;
	if (userData.email) userMetaData.email = userData.email;

	// Add meta data fields with audit trail
	const metaFields = [
		"admin-bca-number",
		"admin-personal-pronouns",
		"admin-personal-year-of-birth",
		"admin-other-club-name",
		"membership_joining_date",
		"billing_address_1",
		"billing_address_2",
		"billing_city",
		"billing_state",
		"billing_postcode",
	];

	metaFields.forEach((field) => {
		if (userData[field] !== undefined) {
			// Add the main field
			userMetaData.meta_data.push({
				key: field,
				value: userData[field],
			});

			// Add audit trail fields
			userMetaData.meta_data.push({
				key: `${field}_marked_given_at`,
				value: datetime,
			});
			userMetaData.meta_data.push({
				key: `${field}_marked_given_by`,
				value: user,
			});
		}
	});

	const encodedAuthInformation = Utilities.base64Encode(
		`${apiusername}:${apipassword}`,
	);

	const options = {
		method: "put",
		contentType: "application/json",
		headers: {
			Authorization: `Basic ${encodedAuthInformation}`,
		},
		payload: JSON.stringify(userMetaData),
		muteHttpExceptions: true,
	};

	const apiurl = `https://www.${apidomain}/wp-json/wc/v3/customers/${userId}`;
	const response = UrlFetchApp.fetch(apiurl, options);

	// Log response for debugging
	console.log("API Response:", response.getContentText());

	// Check response
	if (response.getResponseCode() !== 200) {
		throw new Error(`Failed to update user: ${response.getContentText()}`);
	}

	return response;
}

function pokeToWooUserMeta(data, user_id) {
	return updateWooUser(user_id, data);
}
