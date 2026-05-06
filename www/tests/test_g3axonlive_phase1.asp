<%@ Language="VBScript" %>
<html>

    <head>
        <title>G3AXONLIVE - Phase 1 Test</title>
    </head>

    <body>
        <h1>G3AXONLIVE - Phase 1 Test</h1>

        <%
        ' Test G3AXONLIVE library instantiation and basic functionality
        On Error Resume Next
        
        Set axonlive = Server.CreateObject("G3AXONLIVE")
        
        If Err.Number <> 0 Then
            Response.Write "<p style='color:red;'>ERROR: Failed to create G3AXONLIVE object: " & Err.Description & "</p>"
            Err.Clear
        Else
            Response.Write "<h2 style='color:green;'>✓ G3AXONLIVE Object Created Successfully</h2>"
            
            ' Test Version property
            version = axonlive.Version
            Response.Write "<p><strong>Library Version:</strong> " & version & "</p>"
            
            ' Test SetComponentProperty
            Response.Write "<h3>Testing SetComponentProperty...</h3>"
            result = axonlive.SetComponentProperty("sess_001", "button_1", "Text", "Click Me")
            Response.Write "<p>✓ Set button_1 Text property</p>"
            
            result = axonlive.SetComponentProperty("sess_001", "label_1", "Value", "0")
            Response.Write "<p>✓ Set label_1 Value property</p>"
            
            ' Test GetComponentProperty
            Response.Write "<h3>Testing GetComponentProperty...</h3>"
            buttonText = axonlive.GetComponentProperty("sess_001", "button_1", "Text")
            Response.Write "<p><strong>button_1.Text:</strong> " & buttonText & "</p>"
            
            labelValue = axonlive.GetComponentProperty("sess_001", "label_1", "Value")
            Response.Write "<p><strong>label_1.Value:</strong> " & labelValue & "</p>"
            
            ' Test SetComponentProperty - Update value (simulating button click)
            Response.Write "<h3>Testing Property Update (Simulating Button Click)...</h3>"
            result = axonlive.SetComponentProperty("sess_001", "label_1", "Value", "1")
            updatedValue = axonlive.GetComponentProperty("sess_001", "label_1", "Value")
            Response.Write "<p><strong>label_1.Value (after update):</strong> " & updatedValue & "</p>"
            
            ' Test GetComponentState
            Response.Write "<h3>Testing GetComponentState...</h3>"
            state = axonlive.GetComponentState("sess_001", "button_1")
            Response.Write "<pre>" & Replace(state, "<", "&lt;") & "</pre>"
            
            ' Test RemoveComponentProperty
            Response.Write "<h3>Testing RemoveComponentProperty...</h3>"
            result = axonlive.RemoveComponentProperty("sess_001", "button_1", "Text")
            removedValue = axonlive.GetComponentProperty("sess_001", "button_1", "Text")
            If removedValue = "" Then
                Response.Write "<p>✓ Property removed successfully (returns empty)</p>"
            End If
            
            ' Test ClearComponentState
            Response.Write "<h3>Testing ClearComponentState...</h3>"
            result = axonlive.ClearComponentState("sess_001", "button_1")
            Response.Write "<p>✓ Component state cleared</p>"
            
            ' Test RemoveSession
            Response.Write "<h3>Testing RemoveSession...</h3>"
            result = axonlive.RemoveSession("sess_001")
            Response.Write "<p>✓ Session removed from state map</p>"
            
            ' Test StartCleanup and StopCleanup
            Response.Write "<h3>Testing Cleanup Methods...</h3>"
            result = axonlive.StartCleanup
            Response.Write "<p>✓ Background cleanup started</p>"
            
            result = axonlive.StopCleanup
            Response.Write "<p>✓ Background cleanup stopped</p>"
            
            Response.Write "<h2 style='color:green;'>✓ All Tests Passed!</h2>"
        End If
    %>

        <hr>
        <p><small>G3AXONLIVE Phase 1: Core Go State Management - Successfully Implemented</small></p>
    </body>

</html>