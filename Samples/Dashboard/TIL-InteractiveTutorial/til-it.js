"use strict";

(function () {
  let dashboard;
  let dashboardObjects = [];
  let zoneId = -1;
  let api;
  const tabdocNs = 'tabdoc';
  
  $(document).ready(function () {
    api = tableau.extensions;

    api.initializeAsync().then(async function () {
      dashboard = api.dashboardContent.dashboard;
      dashboardObjects = dashboard.objects;
      console.log(dashboard);

      $("#hide-tutorial").on("click", hideTutorial);
      $("#start-tutorial").on("click", startTutorial);
      $("#show-tutorial").on("click", showTutorial);

      hideTutorial()

      // const { isPresentationMode } = await api.workbook.executeCommandAsync("tabui", "get-is-presentation-mode", {});
      // if (!isPresentationMode) {
      //   $(".container").css("display", "block");

      //   await api.workbook.executeCommandAsync("tabdoc", "set-zone-is-hidden", {
      //     bool: false,
      //     "zone-ids": [zoneId],
      //     dashboard: dashboard.name,
      //   });
      // }
    });
  });

  async function getDashboardZones(name) {
    const dashboardMeta = await api.workbook.executeCommandAsync(tabdocNs, 'get-dashboard', {
      "dashboard": name
    });
    return dashboardMeta?.dashboardPresModel?.zones;
  }

  async function hideTutorial() {

    const zones = await getDashboardZones(dashboard.name);
    console.log("zones", zones);
    
    for (const [key, value] of Object.entries(zones)) {
      console.log('[til-it.js] value', value);
      if(value?.zoneCommon?.zoneType === "web") {
        zoneId = value?.zoneId;
      }
    }
    console.log('[til-it.js] zoneId', zoneId);

    //tabdoc:set-zone-is-hidden bool="true" zone-ids=[803] dashboard="Popular Content"
    await api.workbook.executeCommandAsync(tabdocNs, 'set-zone-is-hidden', {
      "bool": "true",
      "zone-ids": [zoneId],
      "dashboard": dashboard.name
    })

    // tabdoc:edit-web-object-url dashboard="Tutorial" url-string="www.google.com" zone-id="12"
    await api.workbook.executeCommandAsync(tabdocNs, 'edit-web-object-url', {
      "dashboard": dashboard.name,
      "url-string": "http://localhost:8765/Samples/Dashboard/TIL-InteractiveTutorial/blank.html",
      "zone-id": zoneId
    })
    
  }

  async function startTutorial() {

    await api.workbook.executeCommandAsync(tabdocNs, 'edit-web-object-url', {
      "dashboard": dashboard.name,
      "url-string": "http://localhost:8765/Samples/Dashboard/TIL-InteractiveTutorial/top.html",
      "zone-id": zoneId
    })

    await api.workbook.executeCommandAsync(tabdocNs, 'set-zone-is-hidden', {
      "bool": "false",
      "zone-ids": [zoneId],
      "dashboard": dashboard.name
    })

  }

  async function showTutorial() {

    await api.workbook.executeCommandAsync(tabdocNs, 'set-zone-is-hidden', {
      "bool": "false",
      "zone-ids": [zoneId],
      "dashboard": dashboard.name
    })
    
  }






  // async function handleAddSheetButtonClick() {
  //   const sheets = await api.workbook.executeCommandAsync(tabdocNs, 'get-sheets-info', {});
  //   console.log(sheets);
  //   const command = await api.workbook.executeCommandAsync(tabdocNs, 'add-sheet-to-dashboard', {
  //     "add-as-floating": "false",
  //     "dashboard": "Tutorial",
  //     "worksheet": "Which country has the highest avg female height?"
  //   })
  //   console.log(command);
  // }

  /*  addPillToSheet function to convert variables to argument for drop-ui command
   *  @fieldId: example ID [federated.0okqa6q0d7n9uc16bo76i0aq6p3p].[none:ISO:nk]
   *  @destination: "rows" or "columns"
   *  @position: 0-base index of pill order
   *  @targetSheet: name of the target sheet to control
  */

  function addPillToSheet(fieldId, destination, postion, targetSheet) {
    const argsObj = {
      "drag-description" : "",
      "drag-source" : "drag-drop-schema",
      "drop-target" : "drag-drop-shelf",
      "field-encodings" : [
        {
          "fn": fieldId,
          "encoding-type-pres-model":
            {
              "encoding-type": "invalid-encoding",
              "custom-encoding-type-id": ""
            }
        }
      ],
      "is-copy" : "false",
      "is-dead-drop" : "false" ,
      "is-right-drag" : "false",
      "shelf-drag-source-position" : {"is-override": false},
      "shelf-drop-context" : "none",
      "shelf-drop-target-position" : {
        "shelf-type": `${destination}-shelf`,
        "shelf-pos-index": postion,
        "encoding-type-pres-model":
          {
            "encoding-type": "invalid-encoding",
            "custom-encoding-type-id": ""
          },
        "is-override": false
      },
      "target-sheet" : targetSheet
    }
    return argsObj;
  }

  async function handleAddSexToRowsClick() {
    
    const dataSources = await api.workbook.getAllDataSourcesAsync();
    
    const dashboardMeta = await api.workbook.executeCommandAsync(tabdocNs, 'get-dashboard', {
      "dashboard": "Tutorial"
    });

    console.log('[til-it.js] dashboardMeta', dashboardMeta);
    
    let datasource;
    
    dataSources.forEach((ds) => {
      if(ds.name === "NCD_RisC_eLife_2016_height_age18_countries9") {
        datasource = ds;
      }
    });

    const fieldInternalId = `[${datasource.id}].[none:Sex:nk]`;
    console.log('[til-it.js] fieldInternalId', fieldInternalId);

    // Find the zone id of the target sheet
    const zones = dashboardMeta?.dashboardPresModel?.zones;
    let zoneId;
    console.log('[til-it.js] zones', zones);
    
    for (const [key, value] of Object.entries(zones)) {
      if(value?.sheet === "Which country has the highest avg female height?") {
        zoneId = value?.zoneId;
      }
    }
    
    console.log('[til-it.js] zoneId', zoneId);

    
    // tabdoc:set-active-zone skip-if-active="true"   dashboard="Tutorial"   zone-id="8"
    await api.workbook.executeCommandAsync(tabdocNs, 'set-active-zone', {
      "skip-if-active": "true",
      "dashboard": "Tutorial",
      "zone-id": `${zoneId}`
    })

    const command = await api.workbook.executeCommandAsync(tabdocNs, 'drop-on-shelf', addPillToSheet(
      fieldInternalId,
      "rows",
      0,
      "Which country has the highest avg female height?"
    ))

    console.log(command);
  }

      // tabdoc:add-dashboard-object dashboard-object-identifer="object.web" add-as-floating="true" drop-location={"x": 20,"y": 20}  dashboard="Tutorial" 
    // const command = await api.workbook.executeCommandAsync(tabdocNs, 'add-dashboard-object', {
    //   "add-as-floating": "true",
    //   "dashboard": "Tutorial",
    //   "dashboard-object-identifer": "object.web",
    //   "dashboard-object-metadata": { "id": "annotation" },
    //   "drop-location": {"x": 20, "y": 20}
    // })
    // console.log("command", command);

    // { "webZone" : { "url" : "www.google.com" } }

    // tabdoc:is-web-zone-url-valid url-string="DummyUrl"
    // const isValid = await api.workbook.executeCommandAsync(tabdocNs, 'is-web-zone-url-valid', {
    //   "url-string": "DummyUrl"
    // })
    // console.log("isValid", isValid);



})();