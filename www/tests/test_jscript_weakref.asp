<%@ language="JScript" %>
<%
    // Test WeakRef
    var obj1 = { val: 42 };
    var wr = new WeakRef(obj1);
    var target = wr.deref();
    
    Response.Write("WeakRef deref success: " + (target === obj1) + "<br>");
    if (target) {
        Response.Write("WeakRef value: " + target.val + "<br>");
    }
    
    // Test FinalizationRegistry
    var registry = new FinalizationRegistry(function(held) {
        // In this VM, we don't have GC, so this won't run, but the methods should exist.
    });
    
    var obj2 = { name: "test" };
    registry.register(obj2, "some data", obj2);
    
    var unregistered = registry.unregister(obj2);
    Response.Write("FinalizationRegistry unregister success: " + unregistered + "<br>");
    
    var unregisteredAgain = registry.unregister(obj2);
    Response.Write("FinalizationRegistry unregister again: " + unregisteredAgain + "<br>");
%>
