# ForecastSheet

## Overview

This project leverages WeatherAPI and Google Apps Script to fetch and visualize real-time weather updates, detailed summaries, weekly weather predictions, and visualizations of temperature and precipitation. ForecastSheet is designed to make weather data analysis simple, efficient, and accessible directly within your Google Sheets.

## Features

**Real-time Weather Updates**: Fetches the latest weather data for specified locations. <br />
**Weather Summarie**s: Provides comprehensive weather summaries including temperature, condition, humidity, and wind speed. <br />
**Weekly Weather Predictions**: Displays a 7-day weather forecast with detailed information for each day. <br />
**Visualization:** Visualizes temperature and precipitation data for easy analysis. <br />
**Tracking Weather of Multiple Location at the same time**: It allows users to enter multiple location at the same time

## Technologies Used

**APIs**: Utilizes WeatherAPI to fetch real-time and forecast weather data. <br />
**Google Apps Script**: Custom scripts automate data retrieval, processing, and visualization within Google Sheets. <br />

## Getting Started

## Prerequisites

-A Google account <br />
-Access to Google Sheets <br />
-WeatherAPI key (you can get one by signing up on WeatherAPI)<br />

## Installation 

- Clone the Repository<br />
- git clone https://github.com/your-username/ForecastSheet.git <br />
- cd ForecastSheet <br />
- Open Google Sheets <br />
- Open a new or existing Google Sheet.<br />
- Open Script Editor <br />
- Click on Extensions > Apps Script to open the script editor. <br />
- Add the Scripts <br />
- Create two new script files in the script editor and name them Code.gs and Code1.gs. <br />
- Copy the contents of Code.gs from the repository into the Code.gs file in the script editor. <br />
- Copy the contents of Code1.gs from the repository into the Code1.gs file in the script editor. <br />
- Save and Run <br />
- Save the scripts.<br /> 
- Run the onOpen function to add the "Weather" menu to your Google Sheet.<br />

### Usage <br />

Refresh Your Google Sheet <br />
After running the onOpen function, refresh your Google Sheet. You should see a new "Weather" menu in the menu bar. <br />
Fetch Weather Data <br />
Click on the "Weather" menu and select "Weather Forecast." <br />
Enter the city names separated by commas when prompted. <br />
View Weather Data <br />
The weather data for the specified cities will be displayed in new sheets within your Google Sheet. <br />
Each sheet will contain current weather data, a weekly forecast, and a weather summary. <br />

### Try It Out
You can try out ForecastSheet directly in Google Sheets by following this link - https://docs.google.com/spreadsheets/d/1kMdvXcAlX_AMfrzjCVUit1D9ETUQ3SkLTqtt_3Kc2E0/edit?usp=sharing

