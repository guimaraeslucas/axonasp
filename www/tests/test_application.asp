<%@ Language=VBScript %>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>G3pix AxonASP - Application Object Test</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: Tahoma, 'Segoe UI', Geneva, Verdana, sans-serif; padding: 30px; background: #f5f5f5; line-height: 1.6; }
        .container { max-width: 1000px; margin: 0 auto; background: #fff; padding: 30px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        h1 { color: #333; margin-bottom: 10px; border-bottom: 2px solid #667eea; padding-bottom: 10px; }
        h3 { color: #666; margin-top: 15px; margin-bottom: 10px; }
        .intro { background: #e3f2fd; border-left: 4px solid #2196f3; padding: 15px; margin-bottom: 20px; border-radius: 4px; }
        .box { border-left: 4px solid #667eea; padding: 15px; margin-bottom: 15px; background: #f9f9f9; border-radius: 4px; }
        .error { color: #d32f2f; }
        .success { color: #388e3c; }
    </style>
</head>
<body>
    <div class="container">
        <h1>G3pix AxonASP - Application Object Test</h1>
        <div class="intro">
            <p>Tests Application.Lock, Application.Unlock, Application.StaticObjects and variable enumeration.</p>
        </div>

    <div class="box">
        <h3>1. Basic Application Variable Storage</h3>
        <%
            Application("Counter") = 10
            Application("Message") = "Hello from Application"
            Response.Write("Counter: " & Application("Counter") & "<br>")
            Response.Write("Message: " & Application("Message") & "<br>")
        %>
    </div>

    <div class="box">
        <h3>2. Application.Lock and Application.Unlock</h3>
        <%
            Application.Lock
            Application("SafeCounter") = 5
            Response.Write("SafeCounter (locked): " & Application("SafeCounter") & "<br>")
            Application.Unlock
            Response.Write("Unlocked successfully<br>")
        %>
    </div>

    <div class="box">
        <h3>3. Application.StaticObjects - Enumerate All Items</h3>
        <%
            Application("Item1") = "Value1"
            Application("Item2") = "Value2"
            Application("Item3") = "Value3"
            Dim obj
            Response.Write("Iterating through Application.StaticObjects:<br>")
            For Each obj In Application.StaticObjects
                Response.Write("Key: " & obj & " = " & Application(obj) & "<br>")
            Next
        %>
    </div>

    <div class="box">
        <h3>4. Application.StaticObjects Count and Details</h3>
        <%
            Dim count
            count = 0
            Response.Write("All items in Application:<br>")
            Dim key
            For Each key In Application.StaticObjects
                count = count + 1
                Response.Write(count & ". " & key & "<br>")
            Next
            Response.Write("Total items: " & count & "<br>")
        %>
    </div>

    <div class="box">
        <h3>5. Thread-Safety Test (Lock/Unlock)</h3>
        <%
            Application.Lock
            Dim currVal
            currVal = Application("ThreadSafe")
            If IsEmpty(currVal) Or IsNull(currVal) Then
                currVal = 0
            End If
            currVal = currVal + 1
            Application("ThreadSafe") = currVal
            Response.Write("ThreadSafe Counter: " & Application("ThreadSafe") & "<br>")
            Application.Unlock
        %>
    </div>

    <div class="box">
        <h3>6. Mixed Usage (Lock, StaticObjects, and Direct Access)</h3>
        <%
            Application.Lock
            Application("TestA") = 100
            Application("TestB") = 200
            Application("TestC") = 300
            Response.Write("During lock - direct access TestA: " & Application("TestA") & "<br>")
            Response.Write("Items via StaticObjects during lock:<br>")
            For Each itemKey In Application.StaticObjects
                If InStr(itemKey, "Test") > 0 Then
                    Response.Write("  " & itemKey & " = " & Application(itemKey) & "<br>")
                End If
            Next
            Application.Unlock
            Response.Write("Lock released<br>")
        %>
    </div>

    </div>
</body>
</html>
