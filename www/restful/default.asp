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
        <title>AxonASP RESTful Example</title>
        <link rel="stylesheet" href="../css/axonasp.css" />
    </head>
    <body>
        <div class="container">
            <h1>AxonASP RESTful Example</h1>
            <p>
                This page calls RESTful ASP endpoints using resource-oriented
                requests.
            </p>

            <p>
                <button id="btnList" type="button">GET /users</button>
                <button id="btnOne" type="button">GET /users?id=1</button>
                <button id="btnCreate" type="button">POST /users</button>
            </p>

            <p>
                <label for="txtUserName">Name for create:</label>
                <input id="txtUserName" type="text" value="Charlie" />
            </p>

            <h2>Response</h2>
            <pre id="result">No request executed yet.</pre>

            <p><a href="../default.asp">Back to root default page</a></p>
        </div>

        <script>
            (function () {
                var result = document.getElementById("result");
                var btnList = document.getElementById("btnList");
                var btnOne = document.getElementById("btnOne");
                var btnCreate = document.getElementById("btnCreate");
                var txtUserName = document.getElementById("txtUserName");

                function show(title, data) {
                    result.textContent =
                        title + "\n" + JSON.stringify(data, null, 2);
                }

                function showError(title, err) {
                    result.textContent = title + "\n" + String(err);
                }

                btnList.addEventListener("click", function () {
                    fetch("api/users.asp", { method: "GET" })
                        .then(function (res) {
                            return res.json();
                        })
                        .then(function (data) {
                            show("GET /users succeeded", data);
                        })
                        .catch(function (err) {
                            showError("GET /users failed", err);
                        });
                });

                btnOne.addEventListener("click", function () {
                    fetch("api/users.asp?id=1", { method: "GET" })
                        .then(function (res) {
                            return res.json();
                        })
                        .then(function (data) {
                            show("GET /users?id=1 succeeded", data);
                        })
                        .catch(function (err) {
                            showError("GET /users?id=1 failed", err);
                        });
                });

                btnCreate.addEventListener("click", function () {
                    var formBody = new URLSearchParams();
                    formBody.set("name", txtUserName.value || "");

                    fetch("api/users.asp", {
                        method: "POST",
                        headers: {
                            "Content-Type":
                                "application/x-www-form-urlencoded; charset=UTF-8",
                        },
                        body: formBody.toString(),
                    })
                        .then(function (res) {
                            return res.json();
                        })
                        .then(function (data) {
                            show("POST /users succeeded", data);
                        })
                        .catch(function (err) {
                            showError("POST /users failed", err);
                        });
                });
            })();
        </script>
    </body>
</html>
