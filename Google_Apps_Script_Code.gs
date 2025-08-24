function fetchWeatherData() {
  const apiKey = "Your_API_KEY"; // Replace with your OpenWeatherMap API Key
  const cities = [
    "Amaravati", "Visakhapatnam", "Vijayawada", "Guntur", "Nellore",
    "Kurnool", "Kadapa", "Rajahmundry", "Tirupati", "Anantapur",
    "Ongole", "Eluru", "Machilipatnam", "Chittoor", "Vizianagaram",
    "Srikakulam", "Bhimavaram", "Nandyal", "Proddatur", "Tenali",
    "Hindupur", "Adoni", "Gudivada", "Kakinada"
  ];

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const todaySheet = ss.getSheetByName("Today") || ss.insertSheet("Today");
  const next5Sheet = ss.getSheetByName("Next5Days") || ss.insertSheet("Next5Days");

  const todayDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");

  // --- Clear "Today" sheet if old date ---
  if (todaySheet.getLastRow() > 1) {
    const firstDate = todaySheet.getRange(2, 1).getValue();
    const sheetDate = Utilities.formatDate(new Date(firstDate), Session.getScriptTimeZone(), "yyyy-MM-dd");
    if (sheetDate !== todayDate) {
      todaySheet.clear();
    }
  }

  // --- Set headers ---
  if (todaySheet.getLastRow() === 0) {
    todaySheet.appendRow(["Date", "Time", "City", "Temperature (°C)", "Humidity (%)", "Wind Speed (m/s)", 
                          "Rain Chances (%)", "Pressure (hPa)", "Visibility (m)", "UV Index", "Weather Type"]);
  }
  if (next5Sheet.getLastRow() === 0) {
    next5Sheet.appendRow(["Date", "Day", "City", "Avg Temp (°C)", "Min Temp (°C)", "Max Temp (°C)", 
                          "Avg Humidity (%)", "Avg Wind (m/s)", "Avg Rain Chances (%)", "Dominant Weather"]);
  } else {
    next5Sheet.clear();
    next5Sheet.appendRow(["Date", "Day", "City", "Avg Temp (°C)", "Min Temp (°C)", "Max Temp (°C)", 
                          "Avg Humidity (%)", "Avg Wind (m/s)", "Avg Rain Chances (%)", "Dominant Weather"]);
  }

  // --- Fetch and process data ---
  cities.forEach(city => {
    try {
      // Fetch 5-day forecast data
      const forecastUrl = `https://api.openweathermap.org/data/2.5/forecast?q=${city}&appid=${apiKey}&units=metric`;
      const forecastResponse = UrlFetchApp.fetch(forecastUrl);
      const forecastData = JSON.parse(forecastResponse.getContentText());
      
      // Fetch current weather data for additional fields
      const currentUrl = `https://api.openweathermap.org/data/2.5/weather?q=${city}&appid=${apiKey}&units=metric`;
      const currentResponse = UrlFetchApp.fetch(currentUrl);
      const currentData = JSON.parse(currentResponse.getContentText());
      
      // Fetch UV index data (requires separate API call with lat/lon)
      const lat = currentData.coord.lat;
      const lon = currentData.coord.lon;
      const uvUrl = `https://api.openweathermap.org/data/2.5/uvi?lat=${lat}&lon=${lon}&appid=${apiKey}`;
      let uvIndex = "N/A";
      
      try {
        const uvResponse = UrlFetchApp.fetch(uvUrl);
        const uvData = JSON.parse(uvResponse.getContentText());
        uvIndex = uvData.value;
      } catch (e) {
        console.log(`UV data not available for ${city}: ${e.toString()}`);
      }

      // ---- TODAY SHEET (every 3 hours) ----
      forecastData.list.forEach(entry => {
        const dateTime = new Date(entry.dt * 1000);
        const date = Utilities.formatDate(dateTime, Session.getScriptTimeZone(), "yyyy-MM-dd");
        const time = Utilities.formatDate(dateTime, Session.getScriptTimeZone(), "HH:mm");
        
        // Get weather description (capitalize first letter of each word)
        const weatherType = entry.weather[0].description
          .split(' ')
          .map(word => word.charAt(0).toUpperCase() + word.slice(1))
          .join(' ');

        const row = [
          date,
          time,
          city,
          entry.main.temp,
          entry.main.humidity,
          entry.wind.speed,
          Math.round(entry.pop * 100),
          entry.main.pressure,
          entry.visibility,
          uvIndex,
          weatherType
        ];

        if (date === todayDate) {
          todaySheet.appendRow(row);
        }
      });

      // ---- NEXT5DAYS SHEET (daily averages) ----
      const grouped = {};
      forecastData.list.forEach(entry => {
        const date = Utilities.formatDate(new Date(entry.dt * 1000), Session.getScriptTimeZone(), "yyyy-MM-dd");

        if (!grouped[date]) {
          grouped[date] = { 
            temps: [], 
            humidity: [], 
            wind: [], 
            rain: [], 
            weatherTypes: {} 
          };
        }
        grouped[date].temps.push(entry.main.temp);
        grouped[date].humidity.push(entry.main.humidity);
        grouped[date].wind.push(entry.wind.speed);
        grouped[date].rain.push(entry.pop * 100);
        
        // Count weather types for dominant weather
        const weatherMain = entry.weather[0].main;
        grouped[date].weatherTypes[weatherMain] = (grouped[date].weatherTypes[weatherMain] || 0) + 1;
      });

      Object.keys(grouped).forEach(date => {
        const values = grouped[date];
        
        // Get day name from date
        const dayName = getDayName(date);
        
        // Get dominant weather type (most frequent)
        const dominantWeather = getDominantWeather(values.weatherTypes);
        
        next5Sheet.appendRow([
          date,
          dayName,
          city,
          average(values.temps).toFixed(2),
          Math.min(...values.temps).toFixed(2),
          Math.max(...values.temps).toFixed(2),
          average(values.humidity).toFixed(2),
          average(values.wind).toFixed(2),
          average(values.rain).toFixed(2),
          dominantWeather
        ]);
      });
      
      // Add a small delay to avoid hitting API rate limits
      Utilities.sleep(200);
    } catch (e) {
      console.error(`Error fetching data for ${city}: ${e.toString()}`);
    }
  });
  
  // Format the sheets for better readability
  formatSheets();
}

// --- Helper functions ---
function average(arr) {
  return arr.reduce((a, b) => a + b, 0) / arr.length;
}

function getDayName(dateString) {
  const days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
  const date = new Date(dateString);
  return days[date.getDay()];
}

function getDominantWeather(weatherTypes) {
  let dominantType = '';
  let maxCount = 0;
  
  for (const [weatherType, count] of Object.entries(weatherTypes)) {
    if (count > maxCount) {
      maxCount = count;
      dominantType = weatherType;
    }
  }
  
  // Convert to title case (e.g., "CLOUDS" -> "Clouds")
  return dominantType.charAt(0).toUpperCase() + dominantType.slice(1).toLowerCase();
}

function formatSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const todaySheet = ss.getSheetByName("Today");
  const next5Sheet = ss.getSheetByName("Next5Days");
  
  // Format Today sheet
  if (todaySheet.getLastRow() > 1) {
    todaySheet.getRange(1, 1, 1, todaySheet.getLastColumn()).setFontWeight("bold")
      .setBackground("#e6f2ff");
    todaySheet.autoResizeColumns(1, todaySheet.getLastColumn());
  }
  
  // Format Next5Days sheet
  if (next5Sheet.getLastRow() > 1) {
    next5Sheet.getRange(1, 1, 1, next5Sheet.getLastColumn()).setFontWeight("bold")
      .setBackground("#e6f2ff");
    next5Sheet.autoResizeColumns(1, next5Sheet.getLastColumn());
  }
}

// Create a menu item to run the function manually
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Weather Data')
    .addItem('Fetch Weather Data', 'fetchWeatherData')
    .addToUi();
}