<%@ Language="JavaScript" %>
<%
(function() {
    var p = new Proxy(function(a){ return a+1; }, {
        "apply": function(t, thisArg, args) { return t(args[0]) + 10; }
    });
    Response.Write("apply: " + p(1) + "\n");

    var r = Proxy.revocable(function(){}, {});
    Response.Write("revocable: " + (typeof r.proxy) + "\n");

    var p2 = new Proxy(function(){}, {
        "construct": function(t, a) { return { "name": a[0] }; }
    });
    var o = new p2("Alice");
    Response.Write("construct: " + o.name + "\n");
})();
%>
