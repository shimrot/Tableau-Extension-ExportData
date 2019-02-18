$(document).ready(function () {



  tableau.extensions.initializeAsync({'configure': configure}).then(function() {
    if(tableau.extensions.settings.get(buttonLabelKey)) {
      console.log("label",tableau.extensions.settings.get(buttonLabelKey))
      $('#buttonLabel').html(tableau.extensions.settings.get(buttonLabelKey));
    };
    if (tableau.extensions.environment.context == "desktop") {
      //$('#exportBtn').click(exportToWindow);
      $('#exportBtn').click(exportToDownloadPage);
    } else {
      $('#exportBtn').click(exportToExcel);
    }
    tableau.extensions.settings.addEventListener(tableau.TableauEventType.SettingsChanged, (settingsEvent) => {
      //updateExtensionBasedOnSettings(settingsEvent.newSettings)
      var existingSetting = false;
      $('#exportBtn').attr("disabled", "disabled");
      if(tableau.extensions.settings.get(sheetSettingsKey)) {
        var settings = JSON.parse(tableau.extensions.settings.get(sheetSettingsKey));
        for (var i = 0; i < settings.length; i++) {
          if (settings[i].selected) {
            $('#exportBtn').removeAttr("disabled");
            existingSetting = true;
            break;
          }
        }
      }
      if (!existingSetting) {
        configure();
      }
      if(tableau.extensions.settings.get(buttonLabelKey)) {
        console.log("label",tableau.extensions.settings.get(buttonLabelKey))
        $('#buttonLabel').html(tableau.extensions.settings.get(buttonLabelKey));
      };
    });
    console.log("Checing for existing settings");
    if(!tableau.extensions.settings.get(sheetSettingsKey)) {
      console.log("No settings exist. Initialize meta");
      $('#exportBtn').attr("disabled", "disabled");
      func.initializeMeta(function(meta) {
        console.log("Meta built. Saving", meta);
        func.saveSettings(meta, function(settings) {
          console.log("settings saved", settings);
        });
      });
    } else {
      var meta = JSON.parse(tableau.extensions.settings.get(sheetSettingsKey));
      console.log("Settings found", meta);
      $('#exportBtn').removeAttr("disabled");
    }
  });
});


//Get directory of current window
function curDirPath() {
  const location = window.location.href;
  const dirPath = location.substring(0, location.lastIndexOf("/"));
  return dirPath;
}



function configure() {
  const popupUrl = `${curDirPath()}/configure.html`;
  tableau.extensions.ui.displayDialogAsync(popupUrl, 'Payload Message', { height: 500, width: 500 }).then((closePayload) => {

  }).catch((error) => {
    switch(error.errorCode) {
      case tableau.ErrorCodes.DialogClosedByUser:
        console.log("Dialog was closed by user");
        break;
      default:
        console.error(error.message);
    }
  });
}

function exportToWindow() {
  const popupUrl = `${curDirPath()}/summary.html`;
  tableau.extensions.ui.displayDialogAsync(popupUrl, 'Payload Message', { height: 500, width: 800 }).then((closePayload) => {

  }).catch((error) => {
    switch(error.errorCode) {
      case tableau.ErrorCodes.DialogClosedByUser:
        console.log("Dialog was closed by user");
        break;
      default:
        console.error(error.message);
    }
  });
}



var downloadPage;
function exportToDownloadPage() {

  const uid = servfunc.ID();
  const qry = $.param( {uid: uid} );
  var xlblob = {};

  // Must open new window here, before data is sent (because subsequent operations are using promises and 
  // opening windows within promises is not supported for security reasons)
  // https://stackoverflow.com/a/33362850/2736453
  const remotepopupUrl = `${curDirPath()}/download.html?${qry}`;
  window.open(remotepopupUrl);

  buildExcelBlob( function(wb) {
    var wopts = { bookType:'xlsx', bookSST:true, type:'array', ignoreEC:false };
    var wbout = XLSX.write(wb,wopts);
    xlblob = new Blob([wbout],{type:"application/octet-stream"});

    servfunc.sendBlob(uid, xlblob, function() {
      alert("click download button in opened window to get generated excel file");
    });
  });
}



function exportToExcel() {  
  buildExcelBlob( (wb) => {
    // add ignoreEC:false to prevent excel crashes during text to column
    var wopts = { bookType:'xlsx', bookSST:false, type:'array', ignoreEC:false };
    var wbout = XLSX.write(wb,wopts);
    saveAs(new Blob([wbout],{type:"application/octet-stream"}), "export.xlsx");    
  });
}


// krisd: move excel creation to caller (to support extra export to methodss)
// callback receives a blob to save or transfer
function buildExcelBlob(callback) { 
  func.getMeta(function(meta) {
    console.log("Got Meta", meta);
    // func.saveSettings(meta, function(newSettings) {
      // console.log("Saved settings", newSettings);
      var worksheets = tableau.extensions.dashboardContent.dashboard.worksheets;
      var wb = XLSX.utils.book_new();
      var totalSheets = sheetCount = 0;
      var sheetList = [];
      var columnList = [];
      for (var i =0; i < meta.length; i++) {
        if (meta[i].selected) {
          sheetList.push(meta[i].sheetName);
          columnList.push(meta[i].columns);
          totalSheets = totalSheets + 1;
        }
      }
      for (var i = 0; i < worksheets.length; i++) {
        var sheet = worksheets[i];
        if (sheetList.indexOf(sheet.name) > -1) {
          sheet.getSummaryDataAsync({ignoreSelection: true}).then((data) => {
            var headers = [];
            var columns = data.columns;
            var columnMeta = columnList[sheetCount];
            for (var j = 0; j < columnMeta.length; j++) {
              if (columnMeta[j].selected) {
                headers.push(columnMeta[j].name);
              }
            }
            decodeRows(columns, headers, data.data, function(rows) {
              var ws = XLSX.utils.json_to_sheet(rows, {header:headers});
              var sheetname = sheetList[sheetCount];
              sheetCount = sheetCount + 1;
              XLSX.utils.book_append_sheet(wb, ws, sheetname);
              if (sheetCount == totalSheets) {
                callback(wb);
              }
            });
          });
        }
     }
  })
}


// krisd: Remove recursion to work with larger data sets
// and translate cell data types
function decodeRows(columns, headers, dataset, callback) {
  let retArr = [];

  for (let i=0; i<dataset.length; i++) {
    let thisRow = dataset[i];
    let meta = {};
    for (let j = 0; j < columns.length; j++) {
      if (headers.indexOf(columns[j].fieldName) > -1) {
        //meta[columns[j].fieldName] = thisRow[j].formattedValue;

        // krisd: let's assign the sheetjs type according to the summary data column type
        let dtype = undefined;
        let dval = undefined;
        switch (columns[j].dataType) {
          case 'int':
          case 'float': 
            dtype = 'n';
            dval = Number(thisRow[j].value);  // let nums be raw w/o formatting
            if (isNaN(dval)) dval = thisRow[j].formattedValue;  // protect in case issue
            break;
          case 'date':
          case 'date-time': 
            dtype = 'd';
            dval = thisRow[j].formattedValue;
            break;
          case 'bool':
            dtype = 'b';
            dval = thisRow[j].formattedValue;
            break;
          default:
            dtype = 's';
            dval = thisRow[j].formattedValue;
        }

        let o = {v:dval, t:dtype};
        meta[columns[j].fieldName] = o;
      }
    }
    retArr.push(meta);
  }
  callback(retArr);
}
