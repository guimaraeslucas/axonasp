<script runat="server" language="JScript">
    function add(a, b) {
        return this.base + a + b;
    }

    var bound = add.bind({ base: 10 }, 2);
    var obj = {};
    Object.defineProperty(obj, "visible", {
        value: 1,
        writable: true,
        enumerable: true,
        configurable: true,
    });
    Object.defineProperty(obj, "hidden", {
        value: 2,
        writable: false,
        enumerable: false,
        configurable: true,
    });

    var names = Object.getOwnPropertyNames(obj).join(",");
    var desc = Object.getOwnPropertyDescriptor(obj, "hidden");

    Object.preventExtensions(obj);
    obj.extra = 9;

    var sealed = { keep: 5 };
    Object.seal(sealed);
    delete sealed.keep;

    var frozen = { locked: 4 };
    Object.freeze(frozen);
    frozen.locked = 99;

    var iso = new Date(2026, 0, 2, 3, 4, 5).toJSON();

    Response.Write(bound(3) + "|");
    Response.Write(names + "|");
    Response.Write(
        (desc.writable ? "w" : "nw") +
            ":" +
            (desc.enumerable ? "e" : "ne") +
            ":" +
            (desc.configurable ? "c" : "nc") +
            "|"
    );
    Response.Write(
        (Object.isExtensible(obj) ? "ext" : "fixed") +
            ":" +
            (obj.extra === undefined ? "noextra" : "extra") +
            "|"
    );
    Response.Write(
        sealed.keep + ":" + (Object.isSealed(sealed) ? "sealed" : "open") + "|"
    );
    Response.Write(
        frozen.locked +
            ":" +
            (Object.isFrozen(frozen) ? "frozen" : "open") +
            "|"
    );
    Response.Write(iso);
</script>
