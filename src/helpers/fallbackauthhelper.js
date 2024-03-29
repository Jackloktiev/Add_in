/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, location, Office, require */
const documentHelper = require("./documentHelper");
const sso = require("office-addin-sso");
const authModule = require("./authorization");
var loginDialog;

export function dialogFallback() {
  // We fall back to Dialog API for any error.
  const url = "/fallbackauthdialog.html";
  showLoginPopup(url);
}

// This handler responds to the success or failure message that the pop-up dialog receives from the identity provider
// and access token provider.
async function processMessage(arg) {
 
  let messageFromDialog = JSON.parse(arg.message);
  
  //---save access token to the local storage
  window.sessionStorage.setItem('ADtoken', messageFromDialog.result)

  if (messageFromDialog.status === "success") {
    // We now have a valid access token.
    loginDialog.close();
    const response = await sso.makeGraphApiCall(messageFromDialog.result);

    //console.log("Response: ", response);

    //---save user's e-mail and name to the local storage
    window.sessionStorage.setItem('userEmail', response.mail);
    window.sessionStorage.setItem('userDisplayName', response.displayName);
    window.sessionStorage.setItem('userID', response.id);

    //documentHelper.writeDataToOfficeDocument(response);

    authModule.authorizeUser();



  } else {
    // Something went wrong with authentication or the authorization of the web application.
    loginDialog.close();
    sso.showMessage(JSON.stringify(messageFromDialog.error.toString()));
  }
}

// Use the Office dialog API to open a pop-up and display the sign-in page for the identity provider.
function showLoginPopup(url) {
  var fullUrl = location.protocol + "//" + location.hostname + (location.port ? ":" + location.port : "") + url;

  // height and width are percentages of the size of the parent Office application, e.g., PowerPoint, Excel, Word, etc.
  Office.context.ui.displayDialogAsync(fullUrl, { height: 60, width: 30 }, function (result) {
    loginDialog = result.value;
    loginDialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
  });
}
