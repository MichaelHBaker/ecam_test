/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable no-undef */
/* eslint-disable prettier/prettier */

// import fetch from 'node-fetch'; // Using import syntax
// import express from 'express'; // Assuming you have express installed

// const app = express();
// const port = 3001;

// app.get('/weatherdata', async (req, res) => {
//   try {
//     const apiKey = '04eecbfc68f64215926221451240903'; // Your actual API Key
//     const zipCode = '98102'; 
//     const startDate = '2024-03-06'; 

//     const apiUrl = `https://api.weatherapi.com/v1/history.json?q=${zipCode}&dt=${startDate}&key=${apiKey}`;

//     const response = await fetch(apiUrl);

//     if (!response.ok) { 
//        throw new Error(`WeatherAPI error: ${response.status}`);
//     }

//     const data = await response.json();
//     res.json(data); 



//   } catch (error) { 
//     console.error("Error fetching or processing data:", error); 
//     res.status(500).send("Error");  
//   }
// });

// app.listen(port, () => {
//   console.log(`Server listening on port ${port}`);
// });


import fetch from 'node-fetch'; 
import express from 'express';
import mssql from 'mssql'; 

const app = express();
const port = 3001; 

app.get('/weatherdata', async (req, res) => {
  try {
    // WeatherAPI Fetch Logic
    const apiKey = '04eecbfc68f64215926221451240903'; 
    const zipCode = '98102';
    const startDate = '2024-03-06'; 

    const apiUrl = `https://api.weatherapi.com/v1/history.json?q=${zipCode}&dt=${startDate}&key=${apiKey}`;
    const response = await fetch(apiUrl);
    const jsonString = await response.text();
    const weatherData = JSON.parse(jsonString); 

    const maxTempF = weatherData.forecast.forecastday[0].day.maxtemp_f;

    // Success Response (No SQL Insertion here)
    // res.json({ message: 'Data fetched from WeatherAPI' }); 
    res.json(weatherData);

  } catch (error) {
    console.error("Error fetching weather data:", error); 
    res.status(500).send('Error fetching weather data');
  }
});

// New Endpoint for SQL Insertion
app.post('/insertweatherdata', async (req, res) => {
  console.log('Post weather data on server');
  try {
    const { temperature } = req.body;

    // SQL Connection Configuration 
    const sqlConfig = {
      server: 'mikexps13',  
      database: 'ecam', 
      options: {
        trustedConnection: true
      }
    };

    // SQL Insertion Logic
    await mssql.connect(sqlConfig);
    await mssql.query(`INSERT INTO weather (max_temp_f) VALUES (${temperature})`);
    await mssql.close();

    res.json({ message: 'Data inserted into SQL' }); 

  } catch (error) {
    console.error("Error inserting data into SQL Server:", error); 
    res.status(500).send('Error inserting data into SQL Server');
  }
});

app.listen(port, () => {
  console.log(`Server listening on port ${port}`);
});

