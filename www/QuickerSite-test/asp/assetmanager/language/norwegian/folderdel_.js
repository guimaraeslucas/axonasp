function getTxt(s)
	{
    switch(s)
        {
		case "Folder deleted.": return "Mappen er slettet.";
		case "Folder does not exist.": return "Mappen eksisterer ikke.";
		case "Cannot delete Asset Base Folder.": return "Hovedmappen kan ikke slettes.";
        }
    }
function loadTxt()
    {
    document.getElementById("btnCloseAndRefresh").value = "Lukk og oppdater";
    }
function writeTitle()
    {
    document.write("<title>Slett mappe</title>")
    }