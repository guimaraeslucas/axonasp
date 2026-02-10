function getTxt(s)
	{
    switch(s)
        {
		case "Folder deleted.": return "R&eacute;pertoire effac&eacute;.";
		case "Folder does not exist.": return "Le r&eacute;pertoire n'existe pas.";
		case "Cannot delete Asset Base Folder.": return "Impossible d'effacer le r&eacute;pertoire racine.";
        }
    }
function loadTxt()
    {
    document.getElementById("btnCloseAndRefresh").value = "fermer & actualiser";
    }
function writeTitle()
    {
    document.write("<title>Effacer les r\u00E9pertoires</title>")
    }