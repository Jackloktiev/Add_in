/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, require */

const ssoAuthHelper = require("./../helpers/ssoauthhelper");


Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    // Initialize the notification mechanism and hide it
    // var element = document.querySelector(".MessageBanner");
    // messageBanner = new components.MessageBanner(element);
    // messageBanner.hideBanner();
    
    //Add click handler to Authorization button
    document.getElementById("getGraphDataButton").onclick = ssoAuthHelper.getGraphData;

    //Add click handlers to "Load workbook" buttons
    document.getElementById("open-ar-reconciliation-workbook").onclick = OpenARReconciliationWorkbook;
    document.getElementById("open-ap-reconciliation-workbook").onclick = OpenAPReconciliationWorkbook;
    document.getElementById("open-productivity-jc-detail-workbook").onclick = OpenJCDetailsWorkbook;
    document.getElementById("open-productivity-tracker-workbook").onclick = OpenProductivityTrackerWorkbook;

    //Add click handlers to action buttons
    document.getElementById("save-jc-productivity-data").onclick = PushJCProgressEntryIntoVista;
    document.getElementById("load-jc-productivity-data").onclick = LoadJCProgressEntry;

    //---temporary buttons - delete for production
    document.getElementById("local-storage-save").onclick = setLocalStorage;
    document.getElementById("local-storage-get").onclick = getLocalStorage;
    console.log("Initialization complete...");
  }
});

function OpenARReconciliationWorkbook() {

  showNotification('Update', 'Opening AR Reconciliation Workbook. Please wait...');
  const sasUrl = 'https://olsenconsultingaddn.blob.core.windows.net/sharedfiles/AR%20Reconciliation%20Workbook%20V2.xlsm?sp=r&st=2021-11-21T17:33:22Z&se=2023-01-01T01:33:22Z&spr=https&sv=2020-08-04&sr=b&sig=0xMdA8NfPxON4ETOLJ6eoeYiCiA9s57sjOhXwkolCv8%3D'
  LoadWorkbookFromPath(sasUrl, 'AR Reconciliation Workbook Loaded Successfully.');
}

function OpenAPReconciliationWorkbook() {
  showNotification('Update', 'Opening AP Reconciliation Workbook. Please wait...');
  const sasUrl = 'https://olsenconsultingaddn.blob.core.windows.net/sharedfiles/AP%20Reconciliation%20Workbook%20V6.xlsm?sp=r&st=2021-11-21T17:52:32Z&se=2023-01-01T01:52:32Z&spr=https&sv=2020-08-04&sr=b&sig=vrk%2FRNNfoC29qYQ98rIhkl%2FbspKy5eZB2BitinRpQ%2BA%3D';
  LoadWorkbookFromPath(sasUrl, 'AP Reconciliation Workbook Loaded Successfully.');
}

function OpenProductivityTrackerWorkbook() {
  showNotification('Update', 'Opening Productivity Tracker Workbook. Please wait...');
  const sasUrl = 'https://olsenconsultingaddn.blob.core.windows.net/sharedfiles/Productivity%20Tracker%20-%20for%20Sunny.xlsm?sp=r&st=2021-11-21T17:53:27Z&se=2023-01-01T01:53:27Z&spr=https&sv=2020-08-04&sr=b&sig=Nq7wGlQDIXoaI078kMWKLn%2BiaIduILeqb4Ma0VPm9sE%3D';
  LoadWorkbookFromPath(sasUrl, 'Productivity Tracker Workbook Loaded Successfully.');
}

function OpenJCDetailsWorkbook() {
  showNotification('Update', 'Opening JC Detail Workbook. Please wait...');
  const sasUrl = 'https://olsenconsultingaddn.blob.core.windows.net/sharedfiles/JC%20Detail%20for%20T%26M%20Billing.xlsx?sp=r&st=2021-11-21T17:54:02Z&se=2023-01-01T01:54:02Z&spr=https&sv=2020-08-04&sr=b&sig=2TN3%2FZW9smnTRxuDySh7fEJo2eGD4iSNtCuY6NpZcDM%3D';
  LoadWorkbookFromPath(sasUrl, 'JC Detail Workbook Loaded Successfully.');
}

function LoadWorkbookFromPath(path, status_message) {
  console.log('load function started')
  var file = path;
  var request = new XMLHttpRequest();
  request.open("GET", file, true);
  request.responseType = 'blob';
  request.onreadystatechange = function () {
      if (request.readyState === 4) {
          if (request.status === 200 || request.status == 0) {
              console.log(request);
              //var allText = rawFile.response;
              //var allText = document.getElementById('file');
              //console.log(allText);


              //var myFile = document.getElementById("file");
              var reader = new FileReader();
              //console.log(myFile.files[0]);

              reader.onload = (function (event) {
                  Excel.run(function (context) {
                      // Remove the metadata before the base64-encoded string.
                      var startIndex = reader.result.toString().indexOf("base64,");
                      var externalWorkbook = reader.result.toString().substr(startIndex + 7);


                      Excel.createWorkbook(externalWorkbook);
                      showNotification('Update', status_message);
                      return context.sync();
                  }).catch(function (error) {
                      console.log("Error: " + error);
                      if (error instanceof OfficeExtension.Error) {
                          console.log("Debug info: " + JSON.stringify(error.debugInfo));
                      }
                  });
              });

              // Read the file as a data URL so we can parse the base64-encoded string.
              reader.readAsDataURL(request.response);

          }
      }
  }
  request.send(null);
}

function showNotification(header, content) {
  console.log("show notification function");
  // $("#notification-header").text(header);
  // $("#notification-body").text(content);
  // messageBanner.showBanner();
  // messageBanner.toggleExpansion();
};

function PushJCProgressEntryIntoVista() {
  Excel.run(function (context) {
      var sheet = context.workbook.worksheets.getItem("JC Progress Entry");
      var jc_progress_entry = sheet.tables.getItem("JC_Progress_Entry");
      // Get data from the header row
      var headerRange = jc_progress_entry.getHeaderRowRange().load("values");

      // Get data from the table
      var bodyRange = jc_progress_entry.getDataBodyRange().load("values");

      //var bodyValues = bodyRange.values;

      //console.log(headerRange);
      //console.log(bodyValues);

      //for (var i = 0; i < 5; i++) {
      //    console.log(bodyRange.getRow(i));
      //}


      return context.sync()
          .then(function () {
              showNotification('Please wait', 'Process started successfully, please wait.');
              var bodyValues = bodyRange.values;

              console.log(bodyValues);
              var data = [];
              for (var i = 0; i < bodyValues.length; i++) {
                  var row = bodyValues[i];
                  var obj = {};
                  for (var j = 0; j < row.length; j++) {
                      obj['job'] = row[0];
                      obj['phase'] = row[1];
                      console.log(getJsDateFromExcel(row[3]));
                      obj['date'] = getJsDateFromExcel(row[3]);
                      obj['um'] = row[4];
                      obj['newly_completed_units'] = row[15];
                      obj['total_percent_completed'] = row[19];
                  }
                  if (obj['job'] != 'Total')
                      data.push(obj);
              }

              var param = { 'data': data }
              console.log(param);
              $.ajax({
                  type: 'POST',
                  dataType: 'json',
                  data: JSON.stringify(param),
                  headers: {
                      "Content-Type": "application/json"
                  },
                  url: "http://localhost:81/viewpoint/api/v1.0/jcprogressentry",
                  error: function (xhr, status, error) {

                      var err_msg = ''
                      for (var prop in xhr.responseJSON) {
                          err_msg += prop + ': ' + xhr.responseJSON[prop] + '\n';
                      }
                      console.log(err_msg);
                      //alert(err_msg);
                  },
                  success: function (result) {
                      console.log(result);
                      showNotification('Update', 'JC Progress Entry data has been pushed into viewpoint successfully.');
                  }
              });

              return context.sync();
          });


  }).catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}

function LoadJCProgressEntry() {
  var message_notifications_header = "", message_notifications_details = "";
  showNotification('Message:', 'Fetching viewpoint data. Please wait.');
  //Calling viewpoint api for getting data sources of AP Reconciliation
  $.ajax({
      type: 'POST',
      dataType: 'json',
      data: {
          "company": "9"
      },
      url: "http://localhost:81/viewpoint/api/v1.0/loadjcprogressentry",
      error: function (xhr, status, error) {

          var err_msg = ''
          for (var prop in xhr.responseJSON) {
              err_msg += prop + ': ' + xhr.responseJSON[prop] + '\n';
          }

          //alert(err_msg);
      },
      success: function (result) {
          var jc = result.JC;
          console.log(jc);
          showNotification('JC Progress Entry Records', 'Total records returned = ' + jc.length);
          Excel.run(function (context) {
              var sheets = context.workbook.worksheets;
              var sheet = sheets.add("JC");
              sheet.load("name, position");
              sheet.activate();

              var data = [['Co', 'Job', 'PhaseGroup', 'Phase', 'CostType', 'UM', 'ActualUnits', 'ProgressComp', 'ActualDate']];
              for (var i = 0; i < jc.length; i++) {
                  data.push([jc[i]['Co'], jc[i]['Job'], jc[i]['PhaseGroup'], jc[i]['Phase'], jc[i]['CostType'], jc[i]['UM'], jc[i]['ActualUnits'], jc[i]['ProgressComp'], jc[i]['ActualDate']]);
              }

              var range = sheet.getRange("A1").getResizedRange(data.length - 1, data[0].length - 1);
             

              //var sheet = context.workbook.worksheets.getActiveWorksheet();
              //var jcTable = sheet.tables.add("A1:I1", true /*hasHeaders*/);
              //jcTable.name = "JC";

              

              range.values = data;
              //for (var i = 0; i < jc.length; i++) {
              //    jcTable.rows.add(null, /*add rows to the end of the table*/[
              //        [jc[i]['Co'], jc[i]['Job'], jc[i]['PhaseGroup'], jc[i]['Phase'], jc[i]['CostType'], jc[i]['UM'], jc[i]['ActualUnits'], jc[i]['ProgressComp'], jc[i]['ActualDate']]
              //    ]);
              //}

              if (Office.context.requirements.isSetSupported("ExcelApi", "1.2")) {
                  sheet.getUsedRange().format.autofitColumns();
                  sheet.getUsedRange().format.autofitRows();
              }

              
              return context.sync();

          }).catch(function (error) {
              console.log("Error: " + error);
              if (error instanceof OfficeExtension.Error) {
                  console.log("Debug info: " + JSON.stringify(error.debugInfo));
              }
          });

          //alert("Data saved by API successfully.");


      }
  });
}

function setLocalStorage() {
  console.log("button1 clicked...");
  window.sessionStorage.setItem("token", "this is my saved token");
}

function getLocalStorage() {
  console.log("button1 clicked...");
  var token = window.sessionStorage.getItem("ADtoken");
  var userEmail = window.sessionStorage.getItem("userEmail");
  var userDisplayName = window.sessionStorage.getItem("userDisplayName");
  console.log(token);
  console.log(userEmail);
  console.log(userDisplayName);
}
