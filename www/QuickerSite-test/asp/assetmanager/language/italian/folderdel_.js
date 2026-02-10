function getTxt(s)
	{
    switch(s)
        {
		case "Folder deleted.": return "Elimina Cartella.";
		case "Folder does not exist.": return "La cartella non esiste.";
		case "Cannot delete Asset Base Folder.": return "Non puoi eliminare una cartella Risorse.";
        }
    }
function loadTxt()
	{
	document.getElementById("btnCloseAndRefresh").value = "chiudi& rinfresca";
	}
function writeTitle()
	{
	document.write("<title>Elimina Cartella</title>")
	}