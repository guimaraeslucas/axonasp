function getTxt(s)
	{
    switch(s)
        {
		case "Folder deleted.": return "Map verwijderd.";
		case "Folder does not exist.": return "Map bestaat niet.";
		case "Cannot delete Asset Base Folder.": return "Basismap kan niet worden verwijderd.";
        }
    }
function loadTxt()
    {
    document.getElementById("btnCloseAndRefresh").value = "sluiten & vernieuwen";
    }
function writeTitle()
    {
    document.write("<title>Map Verwijderen</title>")
    }