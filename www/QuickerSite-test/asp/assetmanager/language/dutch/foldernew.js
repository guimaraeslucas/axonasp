function getTxt(s)
    {
    switch(s)
        {
		case "Folder already exists.": return "Map bestaat al.";
		case "Folder created.": return "Map aangemaakt.";
        case "Invalid input.":return "Ongeldige tekens.";
        }
    }
function loadTxt()
    {
    document.getElementById("txtLang").innerHTML = "Nieuwe Map";
    document.getElementById("btnCloseAndRefresh").value = "sluiten & vernieuwen";
    document.getElementById("btnCreate").value = "maken";
    }
function writeTitle()
    {
    document.write("<title>Nieuwe Map</title>")
    }
