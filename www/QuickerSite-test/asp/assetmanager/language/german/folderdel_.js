function getTxt(s)
	{
    switch(s)
        {
		case "Folder deleted.": return "Ordner gel&ouml;scht.";
		case "Folder does not exist.": return "Ordner existiert nicht.";
		case "Cannot delete Asset Base Folder.": return "Der Basisordner kann nicht gel&ouml;scht werden.";
        }
    }
function loadTxt()
    {
    document.getElementById("btnCloseAndRefresh").value = "schlie\u00DFen & aktualisieren";
    }
function writeTitle()
    {
    document.write("<title>"+"Ordner l\u00F6schen"+"</title>")
    }