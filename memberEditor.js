const headerMapping = {
	Forenames: "first_name",
	Surname: "last_name",
	"Previous Name": "previous_name",
	"Membership Number": "admin-bca-number",
	"Primary Club Name": "admin-other-club-name",
	"Joining Date": "membership_joining_date",
	"Fee Paid": "fee_paid",
	"Insurance Status": "insurance_status",
	Email: "user_email",
	Gender: "admin-personal-pronouns",
	"Year Of Birth": "admin-personal-year-of-birth",
	"Address 1": "billing_address_1",
	"Address 2": "billing_address_2",
	Town: "billing_city",
	County: "billing_state",
	Postcode: "billing_postcode",
};

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
	// Add debug logging
	console.log("Opening edit dialog for userId:", userId, "row:", row);

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
        <label for="first_name">Forenames:</label>
        <input type="text" id="first_name" name="first_name">
      </div>
      <div class="form-group">
        <label for="last_name">Surname:</label>
        <input type="text" id="last_name" name="last_name">
      </div>
      <div class="form-group">
        <label for="previous_name">Previous Name:</label>
        <input type="text" id="previous_name" name="previous_name">
      </div>
      <div class="form-group">
        <label for="admin-bca-number">Membership Number:</label>
        <input type="text" id="admin-bca-number" name="admin-bca-number">
      </div>
      <div class="form-group">
        <label for="admin-other-club-name">Primary Club Name:</label>
        <input type="text" id="admin-other-club-name" name="admin-other-club-name" 
               placeholder="Leave empty for The Caving Crew">
      </div>
      <div class="form-group">
        <label for="membership_joining_date">Joining Date (DD/MM/YYYY):</label>
        <input type="text" id="membership_joining_date" name="membership_joining_date">
      </div>
      <div class="form-group">
        <label for="admin-personal-pronouns">Gender:</label>
        <select id="admin-personal-pronouns" name="admin-personal-pronouns">
          <option value="m">Male (m)</option>
          <option value="f">Female (f)</option>
          <option value="n">Non-binary (n)</option>
          <option value="an">Another (an)</option>
        </select>
      </div>
      <div class="form-group">
        <label for="admin-personal-year-of-birth">Year of Birth:</label>
        <input type="text" id="admin-personal-year-of-birth" name="admin-personal-year-of-birth">
      </div>
      <div class="form-group">
        <label for="billing_address_1">Address 1:</label>
        <input type="text" id="billing_address_1" name="billing_address_1">
      </div>
      <div class="form-group">
        <label for="billing_address_2">Address 2:</label>
        <input type="text" id="billing_address_2" name="billing_address_2">
      </div>
      <div class="form-group">
        <label for="billing_city">Town:</label>
        <input type="text" id="billing_city" name="billing_city">
      </div>
      <div class="form-group">
        <label for="billing_state">County:</label>
        <input type="text" id="billing_state" name="billing_state">
      </div>
      <div class="form-group">
        <label for="billing_postcode">Postcode:</label>
        <input type="text" id="billing_postcode" name="billing_postcode">
      </div>
      <input type="hidden" id="userId" name="userId" value="${userId}">
      <input type="hidden" id="row" name="row" value="${row}">
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
        new FormData(form).forEach((value, key) => {
          formData[key] = value;
          // Debug log
          console.log('Form data:', key, value);
        });
        
        // Ensure userId is included
        formData.userId = document.getElementById('userId').value;
        console.log('Submitting with userId:', formData.userId);
        
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
    const rowData = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];

    // Create mapping for special cases
    const headerMapping = {
        "Forenames": "first_name",
        "Surname": "last_name",
        "Previous Name": "previous_name",
        "Membership Number": "admin-bca-number",
        "Primary Club Name": "admin-other-club-name",
        "Joining Date": "membership_joining_date",
        "Email": "user_email",
        "Gender": "admin-personal-pronouns",
        "Year Of Birth": "admin-personal-year-of-birth",
        "Address 1": "billing_address_1",
        "Address 2": "billing_address_2",
        "Town": "billing_city",
        "County": "billing_state",
        "Postcode": "billing_postcode"
    };

    const data = {};
    headers.forEach((header, index) => {
        const mappedKey = headerMapping[header];
        if (mappedKey) {
            let value = rowData[index];
            
            // Handle null/undefined values
            if (value === null || value === undefined) {
                value = "";
            }
            
            // Handle date objects
            if (value instanceof Date) {
                // Convert to DD/MM/YYYY format
                value = Utilities.formatDate(value, Session.getScriptTimeZone(), "dd/MM/yyyy");
            }
            
            // Handle numbers
            if (typeof value === 'number') {
                value = value.toString();
            }
            
            data[mappedKey] = value;
        }
    });

    // Ensure all mapped fields exist even if empty
    Object.keys(headerMapping).forEach(header => {
        const mappedKey = headerMapping[header];
        if (!data.hasOwnProperty(mappedKey)) {
            data[mappedKey] = "";
        }
    });

    // Log the data for debugging
    console.log("Mapped data:", data);

    return data;
}

function saveMemberChanges(formData) {
	// Add debug logging
	console.log("Saving changes for user:", formData.userId);
	console.log("Form data received:", formData);

	const userId = formData.userId;
	const row = Number.parseInt(formData.row);

	if (!userId) {
		throw new Error("No user ID provided");
	}

	try {
		// Use the new updateWooUser function
		const response = updateWooUser(userId, {
			first_name: formData.first_name,
			last_name: formData.last_name,
			"admin-personal-pronouns": formData["admin-personal-pronouns"],
			"admin-personal-year-of-birth": formData["admin-personal-year-of-birth"],
			"admin-bca-number": formData["admin-bca-number"],
			"admin-other-club-name": formData["admin-other-club-name"],
			membership_joining_date: formData.membership_joining_date,
			billing_address_1: formData.billing_address_1,
			billing_address_2: formData.billing_address_2,
			billing_city: formData.billing_city,
			billing_state: formData.billing_state,
			billing_postcode: formData.billing_postcode,
		});

		// Update spreadsheet
		const sheet = SpreadsheetApp.getActive().getSheetByName("BCA-CIM-Proforma");
		updateSheetRow(sheet, row, formData);

		return true;
	} catch (error) {
		console.error("Error details:", error);
		throw new Error(`Failed to save changes: ${error.message}`);
	}
}

function updateSheetRow(sheet, row, formData) {
	const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

	const updates = new Map();
	for (const [header, field] of Object.entries(headerMapping)) {
		let value = formData[field] || "";

		// Handle special cases
		if (header === "Primary Club Name" && !value) {
			value = "The Caving Crew";
		}
		if (header === "Insurance Status") {
			value = formData["admin-other-club-name"] ? "AN" : "C";
		}
		if (header === "Fee Paid") {
			value = formData["admin-other-club-name"] ? "£0" : "£20";
		}

		updates.set(header, value);
	}

	updates.forEach((value, header) => {
		const col = findColumnIndex(sheet, header);
		if (col) {
			sheet.getRange(row, col).setValue(value);
		}
	});
}
