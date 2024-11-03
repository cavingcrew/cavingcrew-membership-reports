function updateMembershipStatus(metaKey, metaValue) {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActive().getSheetByName('BCA-CIM-Proforma');
  const currentRow = sheet.getActiveRange().getRowIndex();

  // Validate row selection
  if (!isValidRow(currentRow)) {
    ui.alert('Please select a valid member row');
    return;
  }

  // Get user details
  const userDetails = getUserDetails(sheet, currentRow);
  if (!userDetails.userId) {
    ui.alert('No User ID Found');
    return;
  }

  // Confirm action with user
  if (!confirmAction(ui, userDetails.firstName, userDetails.userId)) {
    return;
  }

  // Update WordPress
  try {
    const success = updateWordPressMembership(userDetails.userId, metaKey, metaValue);
    if (success) {
      updateSheetStatus(sheet, currentRow);
      return metaValue;
    }
  } catch (error) {
    ui.alert('Error', `Failed to update: ${error.message}`, ui.ButtonSet.OK);
    return `ERROR: ${error.message}`;
  }
}

// Helper functions
function isValidRow(row) {
  return row > 1 && row < 100;
}

function getUserDetails(sheet, row) {
  const idCol = findColumnIndex(sheet, 'id');
  const forenamesCol = findColumnIndex(sheet, 'Forenames');

  return {
    userId: sheet.getRange(row, idCol).getValue(),
    firstName: sheet.getRange(row, forenamesCol).getValue()
  };
}

function confirmAction(ui, firstName, userId) {
  const response = ui.alert(
    'Confirm Update',
    `Update membership status for ${firstName}?\nUser ID: ${userId}`,
    ui.ButtonSet.OK_CANCEL
  );
  return response === ui.Button.OK;
}

function updateWordPressMembership(userId, metaKey, metaValue) {
  const { user, datetime } = getCurrentUserAndTime();

  const data = {
    meta_data: [
      { key: metaKey, value: metaValue },
      { key: `${metaKey}_marked_given_at`, value: datetime },
      { key: `${metaKey}_marked_given_by`, value: user }
    ]
  };

  const response = pokeToWooUserMeta(data, userId);
  const userData = JSON.parse(response.getContentText());
  
  return userData.meta_data.some(
    meta => meta.key === metaKey && meta.value === metaValue
  );
}

function updateSheetStatus(sheet, row) {
  const statusCol = findColumnIndex(sheet, 'Status');
  if (statusCol) {
    sheet.getRange(row, statusCol).setValue('Given');
  }
}

// Wrapper functions for menu items
function unMembership() {
  updateMembershipStatus("cc_member", "no");
}

function giveMembership() {
  updateMembershipStatus("cc_member", "yes");
}
