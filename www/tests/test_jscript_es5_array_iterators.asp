<script runat="server" language="JScript">
    var items = [1, 2, 3];
    var seen = "";

    items.forEach(function (v, i) {
        seen += v + ":" + i + ";";
    });

    var every = items.every(function (v) {
        return v < 4;
    });

    var some = items.some(function (v) {
        return v === 2;
    });

    var mapped = items
        .map(function (v) {
            return v * 2;
        })
        .join(",");

    var filtered = items
        .filter(function (v) {
            return v >= 2;
        })
        .join(",");

    var reduced = items.reduce(function (acc, v) {
        return acc + v;
    }, 0);

    var reducedRight = ["a", "b", "c"].reduceRight(function (acc, v) {
        return acc + v;
    }, "");

    Response.Write(
        seen +
            "|" +
            (every ? "T" : "F") +
            (some ? "T" : "F") +
            "|" +
            mapped +
            "|" +
            filtered +
            "|" +
            reduced +
            "|" +
            reducedRight
    );
</script>
