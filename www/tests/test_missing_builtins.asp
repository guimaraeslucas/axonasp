<%@ Language=VBScript %>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Filter, GetObject, Randomize</title>
    <style>
        body { font-family: Tahoma, 'Segoe UI', Geneva, Verdana, sans-serif; background: #f5f5f5; padding: 24px; color: #222; }
        .card { background: #fff; border-radius: 6px; padding: 18px; margin-bottom: 16px; box-shadow: 0 2px 4px rgba(0,0,0,0.08); }
        h1 { margin-bottom: 12px; }
        h2 { margin: 0 0 8px 0; }
        code { background: #eef; padding: 2px 6px; border-radius: 4px; }
    </style>
</head>
<body>
    <h1>Filter, GetObject, Randomize</h1>

    <div class="card">
        <h2>Filter</h2>
        <%
            Dim items, includeRes, excludeRes, textRes
            items = Array("Alpha", "beta", "Gamma", "alphabet soup", "delta", "Echo")
            includeRes = Filter(items, "a")
            excludeRes = Filter(items, "a", False)
            textRes = Filter(items, "AL", True, 1)

            Response.Write "Include default: " & Join(includeRes, ", ") & "<br>"
            Response.Write "Exclude matches: " & Join(excludeRes, ", ") & "<br>"
            Response.Write "Text compare: " & Join(textRes, ", ")
        %>
    </div>

    <div class="card">
        <h2>Randomize & Rnd</h2>
        <%
            Dim rSeedA, rSeedARepeat, rAfterSeed, rHold, rCustomSeed, rCustomHold, rTimerSample
            rSeedA = Rnd(-3)
            rSeedARepeat = Rnd(-3)
            rAfterSeed = Rnd()
            rHold = Rnd(0)
            Randomize 9876
            rCustomSeed = Rnd()
            rCustomHold = Rnd(0)
            Randomize
            rTimerSample = Rnd()

            Response.Write "Rnd(-3) repeatable: " & rSeedA & " / " & rSeedARepeat & "<br>"
            Response.Write "Hold last with Rnd(0): " & rAfterSeed & " / " & rHold & "<br>"
            Response.Write "Custom seed repeat: " & rCustomSeed & " / " & rCustomHold & "<br>"
            Response.Write "Timer seed sample: " & rTimerSample
        %>
    </div>

    <div class="card">
        <h2>GetObject</h2>
        <%
            Dim dictA, dictB, jsonLib, sample
            Set dictA = GetObject("", "Scripting.Dictionary")
            dictA.Add "alpha", 1
            dictA.Add "beta", 2

            Set dictB = GetObject("Scripting.Dictionary")
            dictB.Add "x", "first"

            Set jsonLib = GetObject("", "G3JSON")
            sample = Array("one", "two")

            Response.Write "Dictionary counts: " & dictA.Count & " / " & dictB.Count & "<br>"
            Response.Write "JSON stringify: " & jsonLib.Stringify(sample)
        %>
    </div>
</body>
</html>
