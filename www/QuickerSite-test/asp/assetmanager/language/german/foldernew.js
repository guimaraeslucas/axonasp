function getTxt(s)
    {
    switch(s)
        {
		case "Folder already exists.": return "Der Ordner existiert bereits.";
		case "Folder created.": return "Ordner wurde angelegt.";
        case "Invalid input.":return "Falsche Eingabe";
        }
    }
function loadTxt()
    {
    document.getElementById("txtLang").innerHTML = "Neuer Ordnername";
    document.getElementById("btnCloseAndRefresh").value = "schlie\u00DFen & aktualisieren";
    document.getElementById("btnCreate").value = "erstellen";
    }
function writeTitle()
    {
    document.write("<title>Ordner anlegen</title>")
    }
