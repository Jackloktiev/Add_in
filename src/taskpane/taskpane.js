/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, require */

const ssoAuthHelper = require("./../helpers/ssoauthhelper");

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("getGraphDataButton").onclick = ssoAuthHelper.getGraphData;
    document.getElementById("local-storage-save").onclick = setLocalStorage;
    document.getElementById("local-storage-get").onclick = getLocalStorage;
    console.log("set set set")
  }
});

function setLocalStorage(){
  console.log('button1 clicked...')
  window.sessionStorage.setItem('token', 'this is my saved token')
}

function getLocalStorage(){
  console.log('button1 clicked...');
  var token = window.sessionStorage.getItem('ADtoken');
  var userEmail = window.sessionStorage.getItem('userEmail');
  var userDisplayName = window.sessionStorage.getItem('userDisplayName');
  console.log(token);
  console.log(userEmail);
  console.log(userDisplayName);
}
