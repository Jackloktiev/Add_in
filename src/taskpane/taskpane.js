/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, require */

const ssoAuthHelper = require("./../helpers/ssoauthhelper");

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {

    //Add click handler to Authorization button
    document.getElementById("getGraphDataButton").onclick = ssoAuthHelper.getGraphData;

    //Add click handlers to action buttons
    document.getElementById("save-jc-productivity-data").onclick = PushJCProgressEntryIntoVista;
    document.getElementById("load-jc-productivity-data").onclick = LoadJCProgressEntry;

  }
});



function showNotification(header, content) {
  console.log("show notification function");
  // $("#notification-header").text(header);
  // $("#notification-body").text(content);
  // messageBanner.showBanner();
  // messageBanner.toggleExpansion();
}

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

    return context.sync().then(function () {
      showNotification("Please wait", "Process started successfully, please wait.");
      var bodyValues = bodyRange.values;

      console.log(bodyValues);
      var data = [];
      for (var i = 0; i < bodyValues.length; i++) {
        var row = bodyValues[i];
        var obj = {};
        for (var j = 0; j < row.length; j++) {
          obj["job"] = row[0];
          obj["phase"] = row[1];
          console.log(getJsDateFromExcel(row[3]));
          obj["date"] = getJsDateFromExcel(row[3]);
          obj["um"] = row[4];
          obj["newly_completed_units"] = row[15];
          obj["total_percent_completed"] = row[19];
        }
        if (obj["job"] != "Total") data.push(obj);
      }

      var param = { data: data };
      console.log(param);
      $.ajax({
        type: "POST",
        dataType: "json",
        data: JSON.stringify(param),
        headers: {
          "Content-Type": "application/json",
        },
        url: "http://localhost:81/viewpoint/api/v1.0/jcprogressentry",
        error: function (xhr, status, error) {
          var err_msg = "";
          for (var prop in xhr.responseJSON) {
            err_msg += prop + ": " + xhr.responseJSON[prop] + "\n";
          }
          console.log(err_msg);
          //alert(err_msg);
        },
        success: function (result) {
          console.log(result);
          showNotification("Update", "JC Progress Entry data has been pushed into viewpoint successfully.");
        },
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
  var message_notifications_header = "",
    message_notifications_details = "";
  showNotification("Message:", "Fetching viewpoint data. Please wait.");
  //Calling viewpoint api for getting data sources of AP Reconciliation
  $.ajax({
    type: "POST",
    dataType: "json",
    data: {
      company: "9",
    },
    url: "http://localhost:81/viewpoint/api/v1.0/loadjcprogressentry",
    error: function (xhr, status, error) {
      var err_msg = "";
      for (var prop in xhr.responseJSON) {
        err_msg += prop + ": " + xhr.responseJSON[prop] + "\n";
      }

      //alert(err_msg);
    },
    success: function (result) {
      var jc = result.JC;
      console.log(jc);
      showNotification("JC Progress Entry Records", "Total records returned = " + jc.length);
      Excel.run(function (context) {
        var sheets = context.workbook.worksheets;
        var sheet = sheets.add("JC");
        sheet.load("name, position");
        sheet.activate();

        var data = [
          ["Co", "Job", "PhaseGroup", "Phase", "CostType", "UM", "ActualUnits", "ProgressComp", "ActualDate"],
        ];
        for (var i = 0; i < jc.length; i++) {
          data.push([
            jc[i]["Co"],
            jc[i]["Job"],
            jc[i]["PhaseGroup"],
            jc[i]["Phase"],
            jc[i]["CostType"],
            jc[i]["UM"],
            jc[i]["ActualUnits"],
            jc[i]["ProgressComp"],
            jc[i]["ActualDate"],
          ]);
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
    },
  });
}

