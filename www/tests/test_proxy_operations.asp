<%@ Language="JScript" %>
<%
    Response.Write("<h1>Test Proxy Operations</h1>");

    var errors = [];
    function assertEqual(actual, expected, msg) {
        if (actual !== expected) {
            errors.push(msg + ": expected " + expected + ", got " + actual);
        }
    }

    // 1. has trap
    var pHas = new Proxy({a: 1}, {
        has: function(target, prop) {
            if (prop === 'b') return true;
            return prop in target;
        }
    });
    assertEqual('a' in pHas, true, "has trap: existing property");
    assertEqual('b' in pHas, true, "has trap: simulated property");
    assertEqual('c' in pHas, false, "has trap: non-existing property");

    // 2. deleteProperty trap
    var targetDel = {a: 1, b: 2};
    var pDel = new Proxy(targetDel, {
        deleteProperty: function(t, prop) {
            if (prop === 'a') return false;
            delete t[prop];
            return true;
        }
    });
    
    var res1 = delete pDel.a;
    var res2 = delete pDel.b;
    
    assertEqual(res1, false, "deleteProperty trap: rejected deletion");
    assertEqual(res2, true, "deleteProperty trap: allowed deletion");
    assertEqual('a' in targetDel, true, "deleteProperty trap: target property 'a' should still exist");
    assertEqual('b' in targetDel, false, "deleteProperty trap: target property 'b' should be deleted");

    // 3. ownKeys trap (Object.keys)
    var pOwnKeys = new Proxy({a: 1, b: 2}, {
        ownKeys: function(target) {
            return ['a', 'c'];
        }
    });
    var keys = Object.keys(pOwnKeys);
    assertEqual(keys.join(","), "a", "ownKeys trap with Object.keys");

    // 4. ownKeys trap (for-in)
    var keysForIn = [];
    for (var k in pOwnKeys) {
        keysForIn.push(k);
    }
    assertEqual(keysForIn.join(","), "a", "ownKeys trap with for-in");

    // 5. Proxy.revocable
    var r = Proxy.revocable({a: 1}, {});
    var pRev = r.proxy;
    var val1 = pRev.a;
    r.revoke();
    var val2 = "fail";
    try {
        val2 = pRev.a;
    } catch(e) {
        val2 = "revoked";
    }
    assertEqual(val1, 1, "Proxy.revocable: access before revoke");
    assertEqual(val2, "revoked", "Proxy.revocable: access after revoke");

    if (errors.length === 0) {
        Response.Write("<p style='color: green;'>All Proxy Operations tests passed!</p>");
    } else {
        Response.Write("<p style='color: red;'>Errors occurred:</p><ul>");
        for (var i = 0; i < errors.length; i++) {
            Response.Write("<li>" + errors[i] + "</li>");
        }
        Response.Write("</ul>");
    }
%>
