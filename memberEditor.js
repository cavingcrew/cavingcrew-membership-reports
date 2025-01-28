function editMemberDetails() {
	const ui = SpreadsheetApp.getUi();
	const sheet = SpreadsheetApp.getActiveSheet();
	const currentRow = sheet.getActiveRange().getRowIndex();

	if (currentRow <= 1) {
		ui.alert("Please select a member row");
		return;
	}

	// Only allow editing on BCA-CIM-Proforma sheet
	if (sheet.getName() !== "BCA-CIM-Proforma") {
		ui.alert('Please use this function from the "BCA-CIM-Proforma" sheet');
		return;
	}

	const idCol = findColumnIndex(sheet, "id");
	const userId = sheet.getRange(currentRow, idCol).getValue();

	if (!userId) {
		ui.alert("Could not find user ID");
		return;
	}

	showEditDialog(userId, currentRow);
}

function showEditDialog(userId, row) {
	const html = HtmlService.createHtmlOutput(`
    <style>
      .form-group { margin-bottom: 10px; }
      label { display: block; margin-bottom: 5px; }
      input, select { width: 100%; padding: 5px; }
    </style>
    <form id="memberForm">
      <div class="form-group">
        <label for="forenames">Forenames:</label>
        <input type="text" id="forenames" name="forenames">
      </div>
      <div class="form-group">
        <label for="surname">Surname:</label>
        <input type="text" id="surname" name="surname">
      </div>
      <div class="form-group">
        <label for="membershipNumber">Membership Number:</label>
        <input type="text" id="membershipNumber" name="membershipNumber">
      </div>
      <div class="form-group">
        <label for="primaryClubName">Primary Club Name:</label>
        <input type="text" id="primaryClubName" name="primaryClubName">
      </div>
      <div class="form-group">
        <label for="joiningDate">Joining Date (DD/MM/YYYY):</label>
        <input type="text" id="joiningDate" name="joiningDate">
      </div>
      <div class="form-group">
        <label for="insuranceStatus">Insurance Status:</label>
        <select id="insuranceStatus" name="insuranceStatus">
          <option value="C">BCA insurance via us (C)</option>
          <option value="AN">BCA insurance via other club (AN)</option>
        </select>
      </div>
      <div class="form-group">
        <label for="gender">Gender:</label>
        <select id="gender" name="gender">
          <option value="m">Male (m)</option>
          <option value="f">Female (f)</option>
          <option value="n">Non-binary (n)</option>
          <option value="an">Another (an)</option>
        </select>
      </div>
      <div class="form-group">
        <label for="yearOfBirth">Year of Birth:</label>
        <input type="text" id="yearOfBirth" name="yearOfBirth">
      </div>
      <div class="form-group">
        <label for="address1">Address 1:</label>
        <input type="text" id="address1" name="address1">
      </div>
      <div class="form-group">
        <label for="address2">Address 2:</label>
        <input type="text" id="address2" name="address2">
      </div>
      <div class="form-group">
        <label for="town">Town:</label>
        <input type="text" id="town" name="town">
      </div>
      <div class="form-group">
        <label for="county">County:</label>
        <input type="text" id="county" name="county">
      </div>
      <div class="form-group">
        <label for="postcode">Postcode:</label>
        <input type="text" id="postcode" name="postcode">
      </div>
      <input type="hidden" id="userId" value="${userId}">
      <input type="hidden" id="row" value="${row}">
      <button type="submit" onclick="submitForm(); return false;">Save Changes</button>
    </form>
    <script>
      function submitForm() {
        const form = document.getElementById('memberForm');
        const formData = {};
        new FormData(form).forEach((value, key) => formData[key] = value);
        google.script.run
          .withSuccessHandler(() => google.script.host.close())
          .withFailureHandler(error => alert(error))
          .saveMemberChanges(formData);
      }
      
      // Load current values
      google.script.run
        .withSuccessHandler(loadFormData)
        .getMemberData(${userId}, ${row});
        
      function loadFormData(data) {
        Object.keys(data).forEach(key => {
          const element = document.getElementById(key);
          if (element) element.value = data[key];
        });
      }
    </script>
  `)
		.setWidth(400)
		.setHeight(600);

	SpreadsheetApp.getUi().showModalDialog(html, "Edit Member Details");
}

function getMemberData(userId, row) {
	const sheet = SpreadsheetApp.getActive().getSheetByName("BCA-CIM-Proforma");
	const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
	const rowData = sheet
		.getRange(row, 1, 1, sheet.getLastColumn())
		.getValues()[0];

	const data = {};
	headers.forEach((header, index) => {
		data[header.toLowerCase().replace(/\s+/g, "")] = rowData[index];
	});

	return data;
}

function saveMemberChanges(formData) {
	const userId = formData.userId;
	const row = Number.parseInt(formData.row);

	// Update WordPress
	const userMetaData = {
		meta_data: [
			{ key: "admin-personal-pronouns", value: formData.gender },
			{ key: "admin-personal-year-of-birth", value: formData.yearOfBirth },
			{ key: "billing_address_1", value: formData.address1 },
			{ key: "billing_address_2", value: formData.address2 },
			{ key: "billing_city", value: formData.town },
			{ key: "billing_postcode", value: formData.postcode },
			{ key: "admin-bca-number", value: formData.membershipNumber },
		],
	};

	try {
		const response = pokeToWooUserMeta(userMetaData, userId);
		if (response.getResponseCode() !== 200) {
			throw new Error("Failed to update WordPress user data");
		}

		// Update spreadsheet
		const sheet = SpreadsheetApp.getActive().getSheetByName("BCA-CIM-Proforma");
		updateSheetRow(sheet, row, formData);

		return true;
	} catch (error) {
		throw new Error(`Failed to save changes: ${error.message}`);
	}
}

function updateSheetRow(sheet, row, formData) {
	const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

	const updates = new Map([
		["Forenames", formData.forenames],
		["Surname", formData.surname],
		["Membership Number", formData.membershipNumber],
		["Primary Club Name", formData.primaryClubName],
		["Joining Date", formData.joiningDate],
		["Insurance Status", formData.insuranceStatus],
		["Gender", formData.gender],
		["Year Of Birth", formData.yearOfBirth],
		["Address 1", formData.address1],
		["Address 2", formData.address2],
		["Town", formData.town],
		["County", formData.county],
		["Postcode", formData.postcode],
	]);

	updates.forEach((value, header) => {
		const col = findColumnIndex(sheet, header);
		if (col) {
			sheet.getRange(row, col).setValue(value);
		}
	});
}
