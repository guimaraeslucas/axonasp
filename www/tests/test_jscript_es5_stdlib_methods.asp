<%
@ Language = "JScript"
%>
<script runat="server" language="JScript">
    function add(a, b) {
        return this.base + a + b;
    }

    var bound = add.bind({ base: 10 }, 2);
    var callVal = add.call({ base: 4 }, 1, 2);
    var applyVal = add.apply({ base: 5 }, [1, 2]);

    var s = "AbcDef";
    var sOut =
        s.charAt(1) +
        ":" +
        s.charCodeAt(1) +
        ":" +
        s.substring(1, 4) +
        ":" +
        s.substr(2, 3) +
        ":" +
        s.slice(1, -1) +
        ":" +
        s.concat("X", "Y") +
        ":" +
        "abc123".match(/\d+/)[0] +
        ":" +
        "abc123".search(/\d+/) +
        ":" +
        s.toLowerCase() +
        ":" +
        s.toUpperCase() +
        ":" +
        "abc".localeCompare("abd");

    var arr = [1, 2, 3, 4];
    var arrSlice = arr.slice(1, 3).join("");
    var arrSpliceRemoved = arr.splice(1, 2, 9, 8).join("");
    var arrAfterSplice = arr.join("");
    var arrShift = arr.shift();
    var arrUnshiftLen = arr.unshift(7, 6);
    var arrReverse = arr.reverse().join("");
    var arrSort = ["b", "a", "c"].sort().join("");
    var arrConcat = [1, 2].concat([3, 4], 5).join("");

    Response.Write("FN=" + bound(3) + ":" + callVal + ":" + applyVal + ";");
    Response.Write("STR=" + sOut + ";");
    Response.Write(
        "ARR=" +
            arrSlice +
            ":" +
            arrSpliceRemoved +
            ":" +
            arrAfterSplice +
            ":" +
            arrShift +
            ":" +
            arrUnshiftLen +
            ":" +
            arrReverse +
            ":" +
            arrSort +
            ":" +
            arrConcat +
            ";"
    );
</script>
