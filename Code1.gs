function onOpen(e) {
  addMenu(); 
}

function addMenu() {
  var menu = SpreadsheetApp.getUi().createMenu('Charts');
  menu.addItem('Temperature Chart', 'TemperatureChart');
  menu.addItem('Precipitation Chart', 'PrecipitationChart');
  menu.addToUi(); 
}

function TemperatureChart() {
  AddChart('Temperature');
}

function PrecipitationChart() {
  AddChart('Precipitation');
}

function AddChart(type) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var lastRow = sheet.getLastRow();
  var charts = sheet.getCharts();

  for (var i in charts) {
    var chart = charts[i];
    sheet.removeChart(chart);
  }

  if (type == 'Temperature') {
    var chart = sheet.newChart()
      .setChartType(Charts.ChartType.LINE)
      .setOption('title', 'Temperature Chart')
      .setOption('hAxis.title', 'Date')
      .setOption('vAxis.title', 'Temperature (°C)')
      .addRange(sheet.getRange("D3:E" + lastRow))
      .setPosition(11, 4, 0, 0)
      .build();
  }

  if (type == 'Precipitation') {
    var chart = sheet.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .setOption('title', 'Precipitation Chart')
      .setOption('hAxis.title', 'Date')
      .setOption('vAxis.title', 'Precipitation (mm)')
      .addRange(sheet.getRange("D3:H" + lastRow)) // Assuming precipitation data is in column E
      .setOption('series', {
        //0: { color: 'blue', labelInLegend: 'Condition' },
        0: { color: 'red', labelInLegend: 'Temperature' },
        1: { color: 'yellow', labelInLegend: 'Humidity' },
        2: { color: 'orange', labelInLegend: 'Wind Speed' },
      })
      .setPosition(11, 4, 0, 0)
      .build();
  }

  sheet.insertChart(chart);
}
