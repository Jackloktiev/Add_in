export function authorizeUser() {

  const Http = new XMLHttpRequest();
  const userID = window.sessionStorage.getItem("userID");
  const url = "http://localhost:8000/gettemplate/?file='progress_entry'&userID=" + userID;
  Http.open("GET", url);
  Http.setRequestHeader("Authorization", sessionStorage.getItem("ADtoken"));
  Http.send();

  Http.onreadystatechange = () => {
    if (Http.readyState === 4 && Http.status === 200) {
        let warningMessave = document.getElementById("warningText");
        warningMessave.setAttribute("class","warning hidden");
        warningMessave.innerHTML = "";

      let SASobjects = JSON.parse(Http.response);

      const buttonContainer = document.getElementById("button-container");
      //loop over objects in response and for each object create a button, then append it to the parent div
      for (let SAS of SASobjects) {
        let button = document.createElement("BUTTON");
        button.setAttribute("class", "button");
        button.setAttribute("id", SAS[0]);
        button.setAttribute("data-SAS", SAS[1]);
        // add a function to download a particular file and pass SAS as argument
        button.onclick = function () {
          let particularSAS = this.dataset.sas;
          return getWorkbook(particularSAS);
        };
        let t = document.createTextNode(decodeButtonText(SAS[0]));
        button.appendChild(t);
        buttonContainer.appendChild(button);
        let pushPullContainer = document.getElementById("push-pull-container");
        pushPullContainer.setAttribute("class", "button-container");
      }

      function decodeButtonText(btnName) {
        switch (btnName) {
          case "AR_recon":
            return "Load AR Reconciliation Workbook";
          case "AP_recon":
            return "Load AR Reconciliation Workbook";
          case "JC_detail":
            return "Load JC Detail Workbook";
          case "Prod_tracker":
            return "Load Productivity Tracker Workbook";
          default:
            return "Add btn to a function";
        }
      }

      return;
    }else if(Http.readyState === 4 && Http.status === 600){
        //show message "User not authorizedr"
        let warningMessave = document.getElementById("warningText");
        warningMessave.setAttribute("class","warning");
        warningMessave.innerHTML="User not authorized. Please contact Olsen Consulting "
        // let t = document.createTextNode("User not authorized. Please contact Olsen Consulting ");
        // warningMessave.appendChild(t);
    }else{
        console.log( Http.status)
        //show message "Authentication server error"
        let warningMessave = document.getElementById("warningText");
        warningMessave.setAttribute("class","warning");
        warningMessave.innerHTML = "Authentication server error ";
        // let t = document.createTextNode("Authentication server error ");
        // warningMessave.appendChild(t);
    }
  };
}

function getWorkbook(path) {
  console.log("getWorkbook function started");
  var file = path;
  var request = new XMLHttpRequest();
  request.open("GET", file, true);
  request.responseType = "blob";
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

        reader.onload = function (event) {
          Excel.run(function (context) {
            // Remove the metadata before the base64-encoded string.
            var startIndex = reader.result.toString().indexOf("base64,");
            var externalWorkbook = reader.result.toString().substr(startIndex + 7);

            Excel.createWorkbook(externalWorkbook);
            return context.sync();
          }).catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
              console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
          });
        };

        // Read the file as a data URL so we can parse the base64-encoded string.
        reader.readAsDataURL(request.response);
      }
    }
  };
  request.send(null);
}
