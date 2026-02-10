function getTxt(s)
	{
    switch(s)
        {
		case "Folder deleted.": return "Mappen har tagits bort.";
		case "Folder does not exist.": return "Mappen finns inte.";
		case "Cannot delete Asset Base Folder.": return "Kan inte ta bort huvudmappen.";
        }
    }
function loadTxt()
    {
    document.getElementById("btnCloseAndRefresh").value = "St\u00E4ng och uppdatera";
    }
function writeTitle()
    {
    document.write("<title>Ta bort mapp</title>")
    }