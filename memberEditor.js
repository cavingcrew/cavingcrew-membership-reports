function editMemberDetails() {
	const ui = SpreadsheetApp.getUi();
	const sheet = SpreadsheetApp.getActiveSheet();
	const currentRow = sheet.getActiveRange().getRowIndex();

	// Validate selection
	if (!validateSelection(sheet, currentRow)) {
		return;
	}

	// Get user ID
	const userId = getUserId(sheet, currentRow);
	if (!userId) {
		ui.alert("Could not find user ID");
		return;
	}

	// Show edit dialog
	showEditDialog(userId, currentRow);
}

function validateSelection(sheet, row) {
	const ui = SpreadsheetApp.getUi();

	if (row <= 1) {
		ui.alert("Please select a member row");
		return false;
	}

	if (sheet.getName() !== "BCA-CIM-Proforma") {
		ui.alert('Please use this function from the "BCA-CIM-Proforma" sheet');
		return false;
	}

	return true;
}

function getUserId(sheet, row) {
	const idCol = findColumnIndex(sheet, "id");
	return sheet.getRange(row, idCol).getValue();
}

function showEditDialog(userId, row) {
	const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 20px; }
      .form-group { margin-bottom: 15px; }
      label { display: block; margin-bottom: 5px; font-weight: bold; }
      input, select { 
        width: 100%; 
        padding: 8px;
        border: 1px solid #ddd;
        border-radius: 4px;
        box-sizing: border-box;
      }
      input:focus, select:focus {
        outline: none;
        border-color: #4285f4;
      }
      button {
        background-color: #4285f4;
        color: white;
        padding: 10px 20px;
        border: none;
        border-radius: 4px;
        cursor: pointer;
        width: 100%;
      }
      button:hover {
        background-color: #357abd;
      }
      .error { 
        color: red;
        margin-top: 5px;
        font-size: 0.9em;
      }
      .loading {
        text-align: center;
        padding: 20px;
        display: none;
      }
    </style>
    <div id="loading" class="loading">Loading member data...</div>
    <form id="memberForm" style="display: none">
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
      <button type="submit" onclick="submitForm(); return false;">
        <span id="submitText">Save Changes</span>
        <span id="savingText" style="display: none">Saving...</span>
      </button>
    </form>
    <div id="errorMessage" class="error"></div>
    <script>
      const form = document.getElementById('memberForm');
      const loading = document.getElementById('loading');
      const errorMessage = document.getElementById('errorMessage');
      const submitText = document.getElementById('submitText');
      const savingText = document.getElementById('savingText');

      function showError(message) {
        errorMessage.textContent = message;
        errorMessage.style.display = 'block';
      }

      function clearError() {
        errorMessage.style.display = 'none';
      }

      function submitForm() {
        clearError();
        submitText.style.display = 'none';
        savingText.style.display = 'inline';
        
        const formData = {};
        new FormData(form).forEach((value, key) => formData[key] = value);
        
        google.script.run
          .withSuccessHandler(() => {
            google.script.host.close();
          })
          .withFailureHandler(error => {
            showError(error);
            submitText.style.display = 'inline';
            savingText.style.display = 'none';
          })
          .saveMemberChanges(formData);
      }
      
      // Load current values
      google.script.run
        .withSuccessHandler(data => {
          loading.style.display = 'none';
          form.style.display = 'block';
          
          Object.keys(data).forEach(key => {
            const element = document.getElementById(key);
            if (element) element.value = data[key];
          });
        })
        .withFailureHandler(error => {
          loading.style.display = 'none';
          showError('Failed to load member data: ' + error);
        })
        .getMemberData(${userId}, ${row});
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

	// Create mapping for special cases
	const headerMapping = {
		"Membership Number": "membershipnumber",
		"Primary Club Name": "primaryclubname",
		"Joining Date": "joiningdate",
		"Year Of Birth": "yearofbirth",
	};

	const data = {};
	headers.forEach((header, index) => {
		// Use mapping if exists, otherwise convert header to lowercase and remove spaces
		const key =
			headerMapping[header] || header.toLowerCase().replace(/\s+/g, "");
		data[key] = rowData[index] || ""; // Add empty string fallback
	});

	// Log the data for debugging
	console.log("Mapped data:", data);

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
