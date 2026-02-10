function getTxt(s)
	{
    switch(s)
        {
		case "Folder deleted.": return "Mappen er slettet.";
		case "Folder does not exist.": return "Mappen eksisterer ikke.";
		case "Cannot delete Asset Base Folder.": return "Rodmappen kan ikke slettes.";
        }
    }
function loadTxt()
    {
    document.getElementById("btnCloseAndRefresh").value = "Luk og opdater";
    }
function writeTitle()
    {
    document.write("<title>Slet mappe</title>")
    }
