<%@ Language="JScript" %>
<%
    // Test JavaScript VMENGINE global constant
    Response.Write("JavaScript Engine ID: " + VMENGINE + "<br/>");
    
    // Verify it's the correct string
    if (VMENGINE === "G3pix AxonASP JavaScript Engine") {
        Response.Write("✓ JavaScript VMENGINE constant works correctly<br/>");
    } else {
        Response.Write("✗ JavaScript VMENGINE returned unexpected value: " + VMENGINE + "<br/>");
    }
%>