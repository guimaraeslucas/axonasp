<%@ Language="JScript" %>
<%
/*
 * AxonASP JScript UI Demo
 * This page demonstrates a user interface for the JScript Weather API.
 */
var pageTitle = "AxonASP JScript Weather API";
var serverTime = new Date().toLocaleString();
%>
<!DOCTYPE html>
<html>
<!--
            
            AxonASP Server
            Copyright (C) 2026 G3pix Ltda. All rights reserved.
            
            Developed by Lucas Guimarães - G3pix Ltda
            Contact: https://g3pix.com.br/
            Project URL: https://g3pix.com.br/axonasp
            
            This Source Code Form is subject to the terms of the Mozilla Public
            License, v. 2.0. If a copy of the MPL was not distributed with this
            file, You can obtain one at https://mozilla.org/MPL/2.0/.
            
            Attribution Notice:
            If this software is used in other projects, the name "AxonASP Server"
            must be cited in the documentation or "About" section.
            
            Contribution Policy:
            Modifications to the core source code of AxonASP Server must be
            made available under this same license terms.
            
            -->
    <head>
        <title><%=pageTitle%></title>
        <link rel="stylesheet" type="text/css" href="/css/axonasp.css">
        <style>
            #weather-card {
                margin-top: 20px;
                display: none;
            }

            .forecast-detail {
                margin-bottom: 8px;
                font-size: 1.1em;
            }

            #raw-json {
                background: #f8f8f8;
                padding: 15px;
                border: 1px solid #d0d0d0;
                font-family: 'Courier New', Courier, monospace;
                overflow: auto;
                max-height: 250px;
                margin-top: 10px;
                border-radius: 4px;
                font-size: 0.9em;
                color: #333;
            }

            .main-content {
                padding: 25px;
                max-width: 900px;
                margin: 0 auto;
            }
        </style>
    </head>

    <body>
        <div id="header">
            <div class="logo">
                <%
                var ax
                ax = Server.CreateObject("G3Axon.Functions")
                %>
                <img src="<%= ax.AxGetLogo() %>" alt="AxonASP" width="43" />
            </div>
            <h1>AxonASP JavaScript (JScript) API Demo</h1>
        </div>


        <div id="main-container">
            <div class="main-content">
                <div id="content">
                    <div class="card card-top-blue">
                        <div class="card-header">
                            <h1>Weather Forecast API Explorer</h1>
                        </div>
                        <div class="card-body">
                            <p style="margin-bottom: 20px;">This demonstration uses <strong>AxonASP's ES5 JScript
                                    Engine</strong> to power both the front-end interface and the back-end API. Select a
                                city below to fetch real-time data from the <code>api.asp</code> endpoint.</p>

                            <div class="actions-row"
                                style="background: #f0f4f8; padding: 15px; border-radius: 8px; border: 1px solid #d0dae5;">
                                <label for="location-select" style="margin-right: 10px; font-weight: bold;">Select
                                    Location:</label>
                                <select id="location-select" class="input-sm"
                                    style="width: 250px; padding: 6px; border: 1px solid #808080; border-radius: 4px;">
                                    <option value="">-- Choose a City --</option>
                                    <option value="sao_paulo">São Paulo, Brazil</option>
                                    <option value="new_york">New York, USA</option>
                                    <option value="london">London, UK</option>
                                    <option value="tokyo">Tokyo, Japan</option>
                                </select>
                                <button class="btn btn-primary" onclick="fetchWeather()" style="margin-left: 10px;">
                                    Get Forecast
                                </button>
                            </div>

                            <div id="loading" style="display: none; margin-top: 20px; text-align: center;">
                                <span class="pill pill-primary">Contacting AxonASP JScript Engine...</span>
                            </div>

                            <div id="error-alert" class="alert alert-error" style="display: none; margin-top: 20px;">
                            </div>

                            <div id="weather-card">
                                <div class="card"
                                    style="margin-top: 20px; border: 1px solid #3366cc; box-shadow: 0 4px 8px rgba(0,0,0,0.1);">
                                    <div class="card-header"
                                        style="background: linear-gradient(180deg, #f0f0f0 0%, #e0e0e0 100%);">
                                        <h2 id="location-display" style="margin: 0; color: #003399;">Location</h2>
                                    </div>
                                    <div class="card-body">
                                        <div class="grid-2">
                                            <div>
                                                <div class="forecast-detail">
                                                    <strong>Condition:</strong>
                                                    <span id="w-condition" class="badge badge-success"
                                                        style="padding: 5px 12px; font-size: 0.9em;">-</span>
                                                </div>
                                                <div class="forecast-detail">
                                                    <strong>Temperature:</strong>
                                                    <span id="w-temp"
                                                        style="font-size: 1.5em; font-weight: bold; color: #cc3300;">-</span>
                                                    <span style="font-size: 1.2em; color: #cc3300;">°C</span>
                                                </div>
                                            </div>
                                            <div>
                                                <div class="forecast-detail">
                                                    <strong>Humidity:</strong>
                                                    <span id="w-humidity" style="font-weight: bold;">-</span>%
                                                </div>
                                                <div class="forecast-detail">
                                                    <strong>Wind Speed:</strong>
                                                    <span id="w-wind" style="font-weight: bold;">-</span> km/h
                                                </div>
                                            </div>
                                        </div>

                                        <div
                                            style="margin-top: 25px; border-top: 2px solid #3366cc; padding-top: 15px;">
                                            <h4 style="color: #003399; margin-bottom: 10px;">System JSON Response:</h4>
                                            <div id="raw-json"></div>
                                        </div>
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>

                    <div class="card" style="margin-top: 30px; border-left: 5px solid #ffd700;">
                        <div class="card-header" style="background: #fffbe6;">
                            <h3 style="color: #856404;">Why use JScript in AxonASP?</h3>
                        </div>
                        <div class="card-body">
                            <div class="grid-2">
                                <div>
                                    <ul style="line-height: 1.6;">
                                        <li><strong>Universal Format:</strong> Native <code>JSON</code> object support
                                            for easy data exchange.</li>
                                        <li><strong>Modern Syntax:</strong> Support for <code>try/catch</code>,
                                            <code>anonymous functions</code>, and <code>ES5 features</code>.
                                        </li>
                                        <li><strong>Zero Configuration:</strong> Runs directly on the AxonASP VM
                                            alongside VBScript.</li>
                                    </ul>
                                </div>
                                <div style="background: #f9f9f9; padding: 15px; border-radius: 8px;">
                                    <code style="display: block; font-size: 0.85em; color: #d63384;">
                                    // In api.asp:<br>
                                    var json = Server.CreateObject("G3JSON");<br>
                                    var data = json.LoadFile("data.json");<br>
                                    Response.Write(JSON.stringify(data));
                                </code>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>

        <div id="status-bar">
            <div
                style="padding: 2px 15px; font-size: 0.85em; color: #444; display: flex; justify-content: space-between;">
                <span>Server Time: <%=serverTime%> | </span>
                <span> AxonASP v2 JScript/JavaScript Implementation</span>
            </div>
        </div>

        <script>
            function fetchWeather() {
                var select = document.getElementById('location-select');
                var loc = select.value;
                if (!loc) {
                    alert('Please select a city to see the forecast.');
                    return;
                }

                var loading = document.getElementById('loading');
                var errorAlert = document.getElementById('error-alert');
                var weatherCard = document.getElementById('weather-card');

                loading.style.display = 'block';
                errorAlert.style.display = 'none';
                weatherCard.style.display = 'none';

                // Simulate slight network delay to show AxonASP processing
                setTimeout(function () {
                    fetch('api.asp?location=' + loc)
                        .then(function (response) {
                            if (!response.ok) throw new Error('Network response was not ok: ' + response.statusText);
                            return response.json();
                        })
                        .then(function (data) {
                            loading.style.display = 'none';
                            if (data.success) {
                                document.getElementById('location-display').innerText = data.data.name + ', ' + data.data.country;
                                document.getElementById('w-condition').innerText = data.data.forecast.condition;
                                document.getElementById('w-temp').innerText = data.data.forecast.temperature;
                                document.getElementById('w-humidity').innerText = data.data.forecast.humidity;
                                document.getElementById('w-wind').innerText = data.data.forecast.wind;

                                // Format JSON for display
                                document.getElementById('raw-json').innerText = JSON.stringify(data, null, 4);
                                weatherCard.style.display = 'block';

                                // Visual feedback
                                console.log('Successfully fetched data for ' + loc + ' at ' + data.dateStr);
                            } else {
                                errorAlert.innerText = 'API Error: ' + data.error;
                                errorAlert.style.display = 'block';
                            }
                        })
                        .catch(function (err) {
                            loading.style.display = 'none';
                            errorAlert.innerText = 'Connection Error: ' + err.message;
                            errorAlert.style.display = 'block';
                        });
                }, 300);
            }
        </script>
    </body>

</html>