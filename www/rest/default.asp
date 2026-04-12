<%
@Language = "VBSCRIPT" CodePage = "65001"
%>
<%
Option Explicit
%>
<!DOCTYPE html>
<html lang="en">
    <head>
        <meta charset="UTF-8" />
        <meta name="viewport" content="width=device-width, initial-scale=1.0" />
        <title>AxonASP REST Basic Example</title>
        <link rel="stylesheet" href="../css/axonasp.css" />
    </head>
    <body>
        <div class="container">
            <h1>AxonASP REST Basic Example</h1>
            <p>
                This page calls classic REST-style ASP endpoints using
                JavaScript requests.
            </p>

            <p>
                <button id="btnStatus" type="button">
                    Call Status Endpoint
                </button>
                <button id="btnEcho" type="button">Call Echo Endpoint</button>
            </p>

            <p>
                <label for="txtName">Name for echo:</label>
                <input id="txtName" type="text" value="AxonASP" />
            </p>

            <h2>Response</h2>
            <pre id="result">No request executed yet.</pre>

            <p><a href="../default.asp">Back to root default page</a></p>
        </div>

        <script>
            (function () {
                var result = document.getElementById("result");
                var btnStatus = document.getElementById("btnStatus");
                var btnEcho = document.getElementById("btnEcho");
                var txtName = document.getElementById("txtName");

                function printResponse(title, data) {
                    result.textContent =
                        title + "\n" + JSON.stringify(data, null, 2);
                }

                function printError(title, err) {
                    result.textContent = title + "\n" + String(err);
                }

                btnStatus.addEventListener("click", function () {
                    fetch("api/status.asp", { method: "GET" })
                        .then(function (res) {
                            return res.json();
                        })
                        .then(function (data) {
                            printResponse("Status request succeeded", data);
                        })
                        .catch(function (err) {
                            printError("Status request failed", err);
                        });
                });

                btnEcho.addEventListener("click", function () {
                    var nameValue = encodeURIComponent(
                        txtName.value || "guest"
                    );

                    fetch("api/echo.asp?name=" + nameValue, { method: "GET" })
                        .then(function (res) {
                            return res.json();
                        })
                        .then(function (data) {
                            printResponse("Echo request succeeded", data);
                        })
                        .catch(function (err) {
                            printError("Echo request failed", err);
                        });
                });
            })();
        </script>
    </body>
</html>
