<script runat="server" language="JScript">
    var nested = [
        [1, 2],
        [3, 4],
    ];
    var vbNested = new VBArray(nested);

    var split = "a,b,c".split(",");
    var vb = new VBArray(split);

    var dims = vbNested.dimensions();
    var text = "" + vb;
    var concat = [].concat(vb).join(",");
    Response.Write(dims + "|" + text + "|" + concat);
</script>
