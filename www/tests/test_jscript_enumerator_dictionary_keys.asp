<script runat="server" language="JScript">
    var dict = Server.CreateObject("Scripting.Dictionary");
    dict.Add("one", 1);
    dict.Add("two", 2);

    var e = new Enumerator(dict);
    var out = "";
    while (!e.atEnd()) {
        out += "[" + e.item() + "]";
        e.moveNext();
    }

    Response.Write(out);
</script>
