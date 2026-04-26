<%
@Language = JScript
%>
<script runat="server" language="JScript">
    function replaceAll(s, findText, replaceText) {
        var out = "" + s;
        if (findText == null || findText === "") {
            return out;
        }
        out = out.split(findText).join(replaceText);
        return out;
    }

    function htmlEncode(v) {
        if (v == null) {
            return "";
        }
        var s = "" + v;
        s = replaceAll(s, "&", "&amp;");
        s = replaceAll(s, "<", "&lt;");
        s = replaceAll(s, ">", "&gt;");
        s = replaceAll(s, '"', "&quot;");
        return s;
    }

    function toInt(v) {
        if (v == null || v == "") {
            return 0;
        }
        return v - 0;
    }

    function row(k, v) {
        return (
            "<tr><th>" +
            htmlEncode(k) +
            "</th><td>" +
            htmlEncode(v) +
            "</td></tr>"
        );
    }

    function renderInlineImage() {
        var img = Server.CreateObject("G3Image");
        img.NewContext(520, 180);

        img.SetHexColor("#f2f6ff");
        img.Clear();

        img.SetHexColor("#c8d8ff");
        img.DrawRectangle(10, 10, 500, 160);
        img.Fill();

        img.SetHexColor("#0b4fd9");
        img.SetLineWidth(4);
        img.DrawLine(20, 25, 500, 25);
        img.Stroke();

        img.SetHexColor("#1f6feb");
        img.DrawCircle(95, 95, 42);
        img.Fill();

        img.SetHexColor("#1158c7");
        img.DrawCircle(95, 95, 52);
        img.Stroke();

        img.SetHexColor("#ffd447");
        img.DrawRectangle(180, 52, 295, 85);
        img.Fill();

        img.SetHexColor("#8a5a00");
        img.SetLineWidth(3);
        img.DrawRectangle(180, 52, 295, 85);
        img.Stroke();

        img.SetHexColor("#0e3f96");
        img.DrawLine(200, 120, 455, 80);
        img.Stroke();

        var bytes = img.RenderViaTemp("png");
        if (img.LastError != "") {
            Response.ContentType = "text/plain";
            Response.Write("G3Image error: " + img.LastError);
            Response.End();
            img.Close();
            return;
        }

        Response.ContentType = "image/png";
        Response.AddHeader("Cache-Control", "no-store, max-age=0");
        Response.BinaryWrite(bytes);
        Response.End();
        img.Close();
    }

    var mode = Request.QueryString("mode");
    if (mode == "g3image-inline") {
        renderInlineImage();
    }

    var metodo = Request.ServerVariables("REQUEST_METHOD");
    var nomeEnviado = "";

    nomeEnviado = String(Request.Form("nome"));
</script>
<!DOCTYPE html>
<html>
    <head>
        <meta charset="UTF-8" />
        <title>AxonASP JScript Comprehensive Demo</title>
        <style>
            body {
                font-family: Tahoma, Verdana, Arial, sans-serif;
                margin: 20px;
                background: #f8f8f8;
            }
            .box {
                border: 1px solid #b8b8b8;
                background: #fff;
                padding: 10px;
                margin-bottom: 12px;
            }
            .small {
                color: #555;
                font-size: 12px;
            }
            table {
                border-collapse: collapse;
                width: 100%;
            }
            th,
            td {
                border: 1px solid #d3d3d3;
                text-align: left;
                padding: 6px;
                vertical-align: top;
            }
            th {
                background: #f0f2f7;
            }
            pre {
                background: #f3f3f3;
                border: 1px solid #ddd;
                padding: 8px;
                white-space: pre-wrap;
            }
        </style>
    </head>
    <body>
        <h1>AxonASP JScript Comprehensive Runtime Demo</h1>
        <p>
            This page runs broad server-side JScript coverage for ASP intrinsic
            objects, ASPError, G3Image, G3JSON, and JScript-only behavior.
        </p>

        <div class="box">
            <h2>1) ASP Intrinsic Objects in One Flow</h2>
            <script runat="server" language="JScript">
                var demoName = Request("name");
                if (demoName == null || demoName == "") {
                    demoName = "Guest";
                }
                var incomingMode = Request.QueryString("mode");

                Response.Buffer = true;
                Response.Charset = "utf-8";
                Response.AddHeader("X-AxonASP-JScript-Demo", "enabled");
                Response.Cookies("jscript_demo") = "active";

                Session("jscript_demo_name") = demoName;
                Session("jscript_demo_hits") =
                    toInt(Session("jscript_demo_hits")) + 1;

                Application.Lock();
                Application("jscript_demo_total_hits") =
                    toInt(Application("jscript_demo_total_hits")) + 1;
                Application.Unlock();

                Response.Write("<table>");
                Response.Write(row("Request(name)", demoName));
                Response.Write(row("Request.QueryString(mode)", incomingMode));
                Response.Write(row("Request.TotalBytes", Request.TotalBytes));
                Response.Write(
                    row(
                        "Request.ServerVariables(REQUEST_METHOD)",
                        Request.ServerVariables("REQUEST_METHOD")
                    )
                );
                Response.Write(
                    row(
                        "Request.ServerVariables(URL)",
                        Request.ServerVariables("URL")
                    )
                );
                Response.Write(
                    row(
                        "Session(jscript_demo_name)",
                        Session("jscript_demo_name")
                    )
                );
                Response.Write(
                    row(
                        "Session(jscript_demo_hits)",
                        Session("jscript_demo_hits")
                    )
                );
                Response.Write(row("Session.SessionID", Session.SessionID));
                Response.Write(
                    row(
                        "Application(jscript_demo_total_hits)",
                        Application("jscript_demo_total_hits")
                    )
                );
                Response.Write(row("Server.MapPath(/)", Server.MapPath("/")));
                Response.Write(row("Response.Charset", Response.Charset));
                Response.Write(row("Response.Buffer", Response.Buffer));
                Response.Write("</table>");
            </script>
        </div>

        <div class="box">
            <h2>2) ASPError and Err Object Coverage</h2>
            <script runat="server" language="JScript">
                Server.ClearLastError();

                var badObj = Server.CreateObject(
                    "AxonASP.Invalid.Library.ProgId"
                );
                var aspErr = Server.GetLastError();

                var createResult = "Unexpected object";
                if (badObj == null || badObj == "") {
                    createResult = "Empty (expected)";
                }

                Response.Write("<table>");
                Response.Write(
                    row("Server.CreateObject(invalid)", createResult)
                );
                Response.Write(row("ASPError.Number", aspErr.Number));
                Response.Write(row("ASPError.Source", aspErr.Source));
                Response.Write(row("ASPError.Description", aspErr.Description));
                Response.Write(row("ASPError.Category", aspErr.Category));
                Response.Write(row("ASPError.File", aspErr.File));
                Response.Write(row("ASPError.Line", aspErr.Line));
                Response.Write(row("ASPError.Column", aspErr.Column));
                Response.Write("</table>");

                Err.Clear();
                Response.Write(
                    "<p class='small'>Err.Description after Err.Clear(): " +
                        htmlEncode(Err.Description) +
                        "</p>"
                );
            </script>
        </div>

        <div class="box">
            <h2>3) G3Image Full Pipeline and Inline Dynamic Image</h2>
            <script runat="server" language="JScript">
                var g3img = Server.CreateObject("G3Image");
                var savePath = "temp/jscript_demo_large.png";

                g3img.DefaultFormat = "png";
                g3img.JPGQuality = 92;
                g3img.NewContext(320, 120);
                g3img.SetHexColor("#f0f8ff");
                g3img.Clear();
                g3img.SetHexColor("#2c6bed");
                g3img.DrawRectangle(10, 10, 300, 100);
                g3img.Stroke();
                g3img.SetHexColor("#3d9a40");
                g3img.DrawEllipse(160, 60, 90, 40);
                g3img.FillPreserve();
                g3img.SetHexColor("#1f4a20");
                g3img.Stroke();
                g3img.SetHexColor("#ffffff");
                g3img.DrawLine(20, 20, 300, 100);
                g3img.Stroke();

                var saveOk = g3img.SavePNG(savePath);
                var inlineImageUrl =
                    "test_jscript.asp?mode=g3image-inline&t=" +
                    Session("jscript_demo_hits");

                Response.Write("<table>");
                Response.Write(row("G3Image.HasContext", g3img.HasContext));
                Response.Write(row("G3Image.Width", g3img.Width));
                Response.Write(row("G3Image.Height", g3img.Height));
                Response.Write(
                    row("G3Image.DefaultFormat", g3img.DefaultFormat)
                );
                Response.Write(row("G3Image.JPGQuality", g3img.JPGQuality));
                Response.Write(
                    row("G3Image.SavePNG(temp/jscript_demo_large.png)", saveOk)
                );
                Response.Write(row("G3Image.LastMimeType", g3img.LastMimeType));
                Response.Write(row("G3Image.LastTempFile", g3img.LastTempFile));
                Response.Write(row("G3Image.LastError", g3img.LastError));
                Response.Write("</table>");

                Response.Write(
                    "<p>Inline image rendered by G3Image + Response.BinaryWrite endpoint:</p>"
                );
                Response.Write(
                    "<img src='" +
                        htmlEncode(inlineImageUrl) +
                        "' alt='Dynamic G3Image' style='max-width:100%; border:1px solid #bbb;' />"
                );
                Response.Write(
                    "<p class='small'>Static file output path (mapped by server): " +
                        htmlEncode(savePath) +
                        "</p>"
                );
            </script>
        </div>

        <div class="box">
            <h2>4) G3JSON Broad Coverage</h2>
            <script runat="server" language="JScript">
                var g3json = Server.CreateObject("G3JSON");
                var sourceJson =
                    '{"name":"AxonASP","version":2,"enabled":true,"features":["JScript","VM","ASP"],"meta":{"owner":"G3Pix"}}';
                var parsed = g3json.Parse(sourceJson);
                var parsedRoundtrip = g3json.Stringify(parsed);

                var customObj = g3json.NewObject();
                customObj.Add("language", "JScript");
                customObj.Add("engine", "AxonASP VM");
                customObj.Add(
                    "sessionHits",
                    toInt(Session("jscript_demo_hits"))
                );
                customObj.Add(
                    "applicationHits",
                    toInt(Application("jscript_demo_total_hits"))
                );
                customObj.Add("strictCheck", true);
                var customObjJson = g3json.Stringify(customObj);

                var emptyArr = g3json.NewArray();
                var emptyArrJson = g3json.Stringify(emptyArr);

                var parsedArray = g3json.Parse("[10,20,30,40]");
                var parsedArrayJson = g3json.Stringify(parsedArray);

                Response.Write("<table>");
                Response.Write(row("G3JSON.Parse(source)", "OK"));
                Response.Write(
                    row("G3JSON.Stringify(parsed)", parsedRoundtrip)
                );
                Response.Write(
                    row("G3JSON.NewObject + Add + Stringify", customObjJson)
                );
                Response.Write(
                    row("G3JSON.NewArray + Stringify", emptyArrJson)
                );
                Response.Write(
                    row("G3JSON.Parse(array) + Stringify", parsedArrayJson)
                );
                Response.Write("</table>");

                Response.Write(
                    "<pre>Source JSON:\n" + htmlEncode(sourceJson) + "</pre>"
                );

                // Imprime os números de 0 a 4
                for (var i = 0; i < 5; i++) {
                    // No navegador ou Node.js você usaria console.log(i)
                    // No Classic ASP, seria Response.Write(i + "<br>")
                    Response.Write("Iteração número: " + i + "<br>");
                }
            </script>
        </div>

        <div class="box">
            <h2>5) JScript Features Not Present in VBScript</h2>
            <script runat="server" language="JScript">
                function makeCounter(seed) {
                    var n = seed;
                    return function () {
                        n = n + 1;
                        return n;
                    };
                }

                var counter = makeCounter(10);
                var c1 = counter();
                var c2 = counter();

                var add = function (a, b) {
                    return a + b;
                };
                var anonCall = add(7, 5);

                var iifeValue = (function (base) {
                    return base * 3;
                })(4);

                var strictEq = 5 === "5";
                var looseEq = 5 == "5";
                var undefinedType = typeof notDeclaredIdentifier;

                var sideEffect = 0;
                function touch() {
                    sideEffect = sideEffect + 1;
                    return true;
                }
                var shortCircuit1 = true || touch();
                var shortCircuit2 = false && touch();

                var caughtText = "";
                var finallyHit = false;
                try {
                    throw "custom-jscript-exception";
                } catch (e) {
                    caughtText = e;
                } finally {
                    finallyHit = true;
                }

                Response.Write("<table>");
                Response.Write(row("Closure counter call #1", c1));
                Response.Write(row("Closure counter call #2", c2));
                Response.Write(row("Anonymous function call", anonCall));
                Response.Write(row("IIFE result", iifeValue));
                Response.Write(row("Strict equality 5 === '5'", strictEq));
                Response.Write(row("Loose equality 5 == '5'", looseEq));
                Response.Write(
                    row("typeof undeclared identifier", undefinedType)
                );
                Response.Write(
                    row("Short-circuit true || touch()", shortCircuit1)
                );
                Response.Write(
                    row("Short-circuit false && touch()", shortCircuit2)
                );
                Response.Write(row("touch() side effect count", sideEffect));
                Response.Write(row("try/catch caught value", caughtText));
                Response.Write(row("finally executed", finallyHit));
                Response.Write("</table>");
            </script>
            <script runat="server" language="JScript">
                var x = 10;
                var operacao = "x * 5";
                var resultado = Eval(operacao);

                Response.Write(resultado);

                var now = Date.now();
                Response.Write(now);
            </script>
        </div>

        <div class="box">
            <h2>6) Quick Access URLs</h2>
            <p>
                Main page:
                <a href="test_jscript.asp?name=AxonASP"
                    >test_jscript.asp?name=AxonASP</a
                >
            </p>
            <p>
                Inline image endpoint only:
                <a href="test_jscript.asp?mode=g3image-inline"
                    >test_jscript.asp?mode=g3image-inline</a
                >
            </p>
        </div>
    </body>
</html>
