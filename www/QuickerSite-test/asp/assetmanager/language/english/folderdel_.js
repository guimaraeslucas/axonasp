function getTxt(s)
	{
    switch(s)
        {
		case "Folder deleted.": return "Folder deleted.";
		case "Folder does not exist.": return "Folder does not exist.";
		case "Cannot delete Asset Base Folder.": return "Cannot delete Asset Base Folder.";
        }
    }
function loadTxt()
	{
	document.getElementById("btnCloseAndRefresh").value = "close & refresh";
	}
function writeTitle()
	{
	document.write("<title>Delete Folder</title>")
	}