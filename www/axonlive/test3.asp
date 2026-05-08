<%@ Language="JScript" %>
<%
var AxonLive = Server.CreateObject("G3AXONLIVE");
AxonLive.InitPage();
var sessionID = Session.SessionID;

function renderInput(v) {
    return '<input id="txtName" data-g3al-id="txtName" data-g3al-event="change" data-g3al-event-name="onchange" value="' + Server.HTMLEncode(String(v || "")) + '">';
}

function renderLabel(v) {
    return '<span id="lblName">' + Server.HTMLEncode(String(v || "")) + '</span>';
}

var nameVal = AxonLive.GetComponentProperty(sessionID, "txtName", "val");
if (nameVal === null || nameVal === undefined) nameVal = "";

if (AxonLive.IsAsyncRequest) {
    var compID = AxonLive.EventComponentID;
    var evtName = AxonLive.EventName;

    if (compID === "txtName" && evtName === "onchange") {
        nameVal = AxonLive.GetEventArg("value");
    }

    AxonLive.SetComponentProperty(sessionID, "txtName", "val", String(nameVal));
    AxonLive.RegisterComponent("txtName", renderInput(nameVal));
    AxonLive.RegisterComponent("lblName", renderLabel(nameVal));
    AxonLive.EndAsyncResponse();
}
%>

<!DOCTYPE html>
<html lang="en">

    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>AxonLive Application</title>
        <link rel="stylesheet" href="/css/axonasp.css">
    </head>

    <body>

        <div id="main-container">
            <div id="content">
                <%
                Response.Write(renderInput("Teste"));
                %>
                <%
                 Response.Write(renderLabel(""));
                %>
                <button id="axl_button_3" class="btn btn-primary" data-g3al-id="axl_button_3" data-g3al-event="click"
                    data-g3al-event-name="onclick">Click Me</button>

            </div>
        </div>
        <script src="/axonlive/g3axonlive.js"></script>
        <script>
            G3AxonLive.init();
        </script>
    </body>

</html>