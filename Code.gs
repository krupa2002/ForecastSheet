function onOpen(e) {
  addWeatherMenu(); 
}

function addWeatherMenu() {
  var menu = SpreadsheetApp.getUi().createMenu('Weather');
  menu.addItem('Weather Forecast', 'weatherForecast');
  menu.addToUi(); 
}

function weatherForecast() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // Prompt the user to input multiple cities
  const response = ui.prompt('Enter city names separated by commas (e.g., City1, City2)', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) {
    return; // User canceled the input
  }

  const citiesInput = response.getResponseText();
  const cities = citiesInput.split(',').map(city => city.trim());

  // Loop through each city
  cities.forEach(city => {
    const sheetName = city.replace(/\s+/g, '_'); // Replace spaces with underscores for tab name
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    } else {
      sheet.clear(); // Clear existing data on the sheet
    }

    // Fetch weather data for the current city
    const apiKey = "YOUR_API_KEY"; // Your WeatherAPI API key
    const apiCall = "YOUR_API_CALL_LINK" +
      encodeURIComponent(city) + // Make sure to encode the city name
      "&key=" +
      apiKey +
      "&days=7"; // Fetch forecast for the next 7 days

    const response = UrlFetchApp.fetch(apiCall, { muteHttpExceptions: true });
    const data = JSON.parse(response.getContentText());

    if (data.error) {
      sheet.getRange(2, 1).setValue("Error: " + data.error.message);
      return;
    }

    const headerPrediction = ["Next Week Weather Prediction"];
    const headerPredictionRange = sheet.getRange(1, 4, 1, headerPrediction.length);
    headerPredictionRange.setValues([headerPrediction]);
    headerPredictionRange.setFontWeight('bold');
    headerPredictionRange.setFontColor('#FFFFFF');
    headerPredictionRange.setBackground('#4CAF50');
    headerPredictionRange.setFontSize(14);

    const forecastDays = data.forecast.forecastday;
    const dates = forecastDays.map(day => day.date);
    const temperatures = forecastDays.map(day => day.day.avgtemp_c);
    const conditions = forecastDays.map(day => day.day.condition.text);
    const humidities = forecastDays.map(day => day.day.avghumidity);
    const windSpeeds = forecastDays.map(day => day.day.maxwind_kph);

    // Set headers
    const headers = ["Date", "Temperature (°C)", "Condition", "Humidity (%)", "Wind Speed (kph)"];
    const headerRange = sheet.getRange(2, 4, 1, headers.length);
    headerRange.setValues([headers]);
    headerRange.setFontWeight('bold');
    headerRange.setFontColor('#FFFFFF');
    headerRange.setBackground('#4CAF50');
    headerRange.setFontSize(14);

    // Populate weather data
    for (let i = 0; i < forecastDays.length; i++) {
      const row = [dates[i], temperatures[i], conditions[i], humidities[i], windSpeeds[i]];
      const rowRange = sheet.getRange(i + 3, 4, 1, row.length);
      rowRange.setValues([row]);
      rowRange.setFontWeight('bold');
      rowRange.setFontColor('#FFFFFF');
      rowRange.setBackground('#FFA500');
      rowRange.setFontSize(14);
    }

    // Display textual weather data for the current city
    const headerCurrent = ["Current Weather Data"];
    const headerCurrentRange = sheet.getRange(1, 1, 1, headerCurrent.length);
    headerCurrentRange.setValues([headerCurrent]);
    headerCurrentRange.setFontWeight('bold');
    headerCurrentRange.setFontColor('#FFFFFF');
    headerCurrentRange.setBackground('#4CAF50');
    headerCurrentRange.setFontSize(14);

    const weatherData = [
      ["Location:", data.location.name],
      ["Country:", data.location.country],
      ["Weather:", forecastDays[0].day.condition.text], // Weather for the first day
      ["Temperature:", forecastDays[0].day.avgtemp_c], // Temperature for the first day
      ["Feels Like:", forecastDays[0].day.avgtemp_c], // Assuming feels like temperature is the same as average temperature for the first day
      ["Wind Speed:", forecastDays[0].day.maxwind_kph], // Max wind speed for the first day
    ];

    const weatherDataRange = sheet.getRange(2, 1, weatherData.length, weatherData[0].length);
    weatherDataRange.setValues(weatherData);
    weatherDataRange.setFontWeight('bold');
    weatherDataRange.setFontColor('#FFFFFF');
    weatherDataRange.setBackground('#FFA500');
    weatherDataRange.setFontSize(14);

    // Add a weather report
    const generateWeatherReport = (location, currentWeather, avgTemperature, windSpeed) => {
      let report = `Weather Report for ${location}:\n\n`;
      
      // Add basic weather data
      report += `- Current Weather: ${currentWeather}\n`;
      report += `- Average Temperature: ${avgTemperature}°C\n`;
      report += `- Wind Speed: ${windSpeed} km/h\n\n`;

      // Add suggestions based on weather conditions
      if (avgTemperature > 35) {
        report += "- Excessive Heat Warning: It's excessively hot. Take precautions to stay cool and hydrated.\n";
      } else if (avgTemperature < 0) {
        report += "- Cold Warning: It's very cold. Bundle up and avoid prolonged exposure to the cold.\n";
      } else {
        // Calculate the temperature range
        const temperatureRange = Math.floor(avgTemperature / 10) * 10;
        
        // Assign messages based on temperature range
        switch (temperatureRange) {
          case 0:
            report += "- Frigid Conditions: Extreme cold. Stay indoors if possible.\n";
            break;
          case 10:
            report += "- Very Cold: Cold weather. Dress warmly if going outside.\n";
            break;
          case 20:
            report += "- Cold: Cool weather. A light jacket may be needed.\n";
            break;
          case 30:
            report += "- Moderate: Enjoy outdoor activities.\n";
            break;
          case 40:
            report += "- Warm: It's warm outside. Dress comfortably.\n";
            break;
          case 50:
            report += "- Hot: It's hot. Stay hydrated and seek shade if necessary.\n";
            break;
          default:
            report += "- The weather is variable. Check for local advisories for more information.\n";
            break;
        }
      }

      return report;
    };

    const headerSummary = ["Weather Summary"];
    const headerSummaryRange = sheet.getRange(10, 1, 1, headerSummary.length);
    headerSummaryRange.setValues([headerSummary]);
    headerSummaryRange.setFontWeight('bold');
    headerSummaryRange.setFontColor('#FFFFFF');
    headerSummaryRange.setBackground('#4CAF50');
    headerSummaryRange.setFontSize(14);

    const report = generateWeatherReport(
      data.location.name,
      forecastDays[0].day.condition.text,
      forecastDays[0].day.avgtemp_c,
      forecastDays[0].day.maxwind_kph
    );

    const reportLines = report.split('\n');
    const reportRange = sheet.getRange(11, 1, reportLines.length, 1); // Ensure the range only spans one column
    reportRange.setValues(reportLines.map(line => [line]));
    reportRange.setFontWeight('bold');
    reportRange.setFontColor('#FFFFFF');
    reportRange.setBackground('#FFA500');
    reportRange.setFontSize(14);

    // Make B1 green
    const b1Range = sheet.getRange('B1');
    b1Range.setBackground('#4CAF50');
    b1Range.setFontWeight('bold');
    b1Range.setFontColor('#FFFFFF');
    b1Range.setFontSize(14);
  });
}
