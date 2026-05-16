<%@ Language="JScript" %>
<%
/*
 * AxonASP Server
 * SharedArrayBuffer Test
 */

function test() {
    Response.Write("--- SharedArrayBuffer Test ---\n");

    // 1. Creation
    var sab = new SharedArrayBuffer(16);
    Response.Write("sab.byteLength: " + sab.byteLength + " (expected: 16)\n");

    // 2. Views
    var u8 = new Uint8Array(sab);
    u8[0] = 123;
    Response.Write("u8[0]: " + u8[0] + " (expected: 123)\n");

    var dv = new DataView(sab);
    dv.setUint32(4, 0xdeadbeef, true);
    Response.Write("dv.getUint32(4): 0x" + dv.getUint32(4, true).toString(16) + " (expected: deadbeef)\n");

    // 3. Slice
    var sliced = sab.slice(4, 8);
    Response.Write("sliced.byteLength: " + sliced.byteLength + " (expected: 4)\n");
    var dv2 = new DataView(sliced);
    Response.Write("dv2.getUint32(0): 0x" + dv2.getUint32(0, true).toString(16) + " (expected: deadbeef)\n");

    // 4. isView
    Response.Write("ArrayBuffer.isView(u8): " + ArrayBuffer.isView(u8) + " (expected: true)\n");
    Response.Write("ArrayBuffer.isView(dv): " + ArrayBuffer.isView(dv) + " (expected: true)\n");

    Response.Write("--- End Test ---\n");
}

test();
%>
