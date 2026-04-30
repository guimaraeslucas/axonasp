<%@ Language="JScript" %>
<%
/*
 * AxonASP JScript API Demo
 * This file demonstrates how to create a JSON API using JScript in AxonASP.
 */

// Set the response type to JSON
Response.ContentType = "application/json";

function handleRequest() {
    try {
        // Demonstrate getting parameters from QueryString or Form
        var locationId = String(Request.QueryString("location"));
        
        // Use the native G3JSON library for efficient file loading
        var jsonLib = Server.CreateObject("G3JSON");
        
        // Load the JSON data file
        // Note: MapPath is handled internally by G3JSON.LoadFile if we provide a relative path
        var data = jsonLib.LoadFile("weather_data.json");
        
        if (!data || !data.locations) {
            return {
                success: false,
                error: "Weather data source not found or invalid."
            };
        }

        var locations = data.locations;
        var result = null;

        // Demonstrate JScript Array-like iteration on VTArray
        if (locationId && locationId !== "undefined" && locationId !== "null" && locationId !== "") {
            for (var i = 0; i < locations.length; i++) {
                var loc = locations[i];
                if (loc.id === locationId) {
                    result = loc;
                    break;
                }
            }
        } else {
            // If no location specified, return all locations (briefly)
            result = locations;
        }

        if (result) {
            return {
                success: true,
                timestamp: new Date().getTime(),
                dateStr: new Date().toLocaleString(),
                data: result,
                serverInfo: {
                    engine: "AxonASP",
                    language: "JScript (ES5)"
                }
            };
        } else {
            return {
                success: false,
                error: "Location '" + locationId + "' not found in database."
            };
        }

    } catch (e) {
        return {
            success: false,
            error: "Internal Server Error: " + e.message
        };
    }
}

// Execute and write result
var apiResult = handleRequest();
Response.Write(JSON.stringify(apiResult));
%>
