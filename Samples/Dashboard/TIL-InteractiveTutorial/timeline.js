const timeline = [
  {
    "timestamp": 5000,
    "add": [
      {
        "func": "addSheetToDashboard",
        "args": {
          "sourceSheet": "Which country has the highest avg female height?",
          "targetDashboard": "Tutorial"
        }
      }
    ],
    "remove" : []
  },
  {
    "timestamp": 10000,
    "add": [
      {
        "func": "addPillToShelf",
        "args": {
          "targetDashboard": "Tutorial",
          "targetSheet": "Which country has the highest avg female height?",
          "datasourceName": "NCD_RisC_eLife_2016_height_age18_countries9",
          "pillName": "Sex",
          "shelf": "rows",
          "position": 0
        }
      }
    ],
    "remove" : []
  }
];