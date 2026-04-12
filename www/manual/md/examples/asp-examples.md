# ASP Example Applications

## Overview

AxonASP ships with four example web applications located under the `www/` directory. Each example demonstrates a common web application pattern implemented in Classic ASP using AxonASP features. The examples are intended as learning references and starting points for new projects.

Access the examples from the web root while the HTTP server is running at `http://localhost:8801/`.

---

## MVC — Model-View-Controller

**Path:** `./www/mvc/`
**URL:** `http://localhost:8801/mvc/`

This example implements the Model-View-Controller pattern in Classic ASP. Responsibilities are separated into three layers:

| Directory / File | Role |
|-----------------|------|
| `mvc/model/site_model.asp` | Data and business logic functions |
| `mvc/controller/site_controller.asp` | Request handling, page routing, assembling data for views |
| `mvc/view/` | HTML generation functions, one per template |
| `mvc/default.asp` | Front controller entry point |

**How it works:** The front controller (`default.asp`) reads a `page` query string parameter, delegates to the appropriate controller function, and the controller calls the model and a view function to produce HTML output.

**Example request:**
```
http://localhost:8801/mvc/?page=home
http://localhost:8801/mvc/?page=example
```

---

## MVVM — Model-View-ViewModel

**Path:** `./www/mvvm/`
**URL:** `http://localhost:8801/mvvm/`

This example implements the Model-View-ViewModel pattern in Classic ASP. The ViewModel layer prepares display-ready data and exposes it to the View, decoupling rendering from business logic.

| Directory / File | Role |
|-----------------|------|
| `mvvm/model/site_model.asp` | Data access and domain logic |
| `mvvm/viewmodel/site_viewmodel.asp` | Data transformation and preparation for display |
| `mvvm/view/` | HTML templates consuming ViewModel output |
| `mvvm/default.asp` | Entry point and routing |

**How it works:** The entry point includes the model and viewmodel layers, resolves the requested route, calls the appropriate ViewModel function, and outputs the rendered HTML.

**Example request:**
```
http://localhost:8801/mvvm/?page=home
http://localhost:8801/mvvm/?page=example
```

---

## REST — Basic REST-Style Endpoints

**Path:** `./www/rest/`
**URL:** `http://localhost:8801/rest/`

This example demonstrates basic REST-style HTTP endpoints using separate `.asp` files under an `api/` subdirectory. A browser front-end page makes JavaScript `fetch` requests to the endpoints and displays the JSON responses.

| Path | Description |
|------|-------------|
| `rest/default.asp` | Demo front-end page with buttons to call the API endpoints |
| `rest/api/status.asp` | Returns a JSON status response |
| `rest/api/echo.asp` | Returns the `name` query string parameter as a JSON echo response |

**How it works:** Each endpoint file in `api/` reads the request, builds a response dictionary using `G3JSON`, and writes the JSON string with the correct `Content-Type` header.

**Example requests:**
```
GET http://localhost:8801/rest/api/status.asp
GET http://localhost:8801/rest/api/echo.asp?name=AxonASP
```

---

## RESTful — Full RESTful API

**Path:** `./www/restful/`
**URL:** `http://localhost:8801/restful/`

This example extends the REST example with a more complete RESTful design, including support for multiple HTTP verbs and resource-oriented endpoints.

| Path | Description |
|------|-------------|
| `restful/default.asp` | Front-end page demonstrating API calls |
| `restful/api/` | API endpoint files organized by resource |

**How it works:** Endpoint scripts branch on `Request.ServerVariables("REQUEST_METHOD")` to handle `GET`, `POST`, `PUT`, and `DELETE` verbs. Responses are JSON formatted using `G3JSON`.

---

## Shared Patterns

All four examples follow these AxonASP Classic ASP conventions:

- `Option Explicit` is declared at the top of each file.
- Objects use `Set ... = Server.CreateObject(...)` syntax.
- JSON output uses `G3JSON` (`Server.CreateObject("G3JSON")`).
- CSS uses `../css/axonasp.css` from the shared stylesheet.
- No framework dependencies — pure Classic ASP and VBScript.

## Remarks

- The example applications are for development and educational use only. They are not hardened for production deployment.
- To adapt an example as a starting point for a real project, copy the directory to a new location under `www/` and modify the files as needed.
- All four examples run without any database or external service. They can be exercised immediately after starting the AxonASP HTTP server.
