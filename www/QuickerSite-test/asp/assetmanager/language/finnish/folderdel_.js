function getTxt(s)
	{
    switch(s)
        {
		case "Folder deleted.": return "Kansio poistettu."";
		case "Folder does not exist.": return "Kansiota ei ole olemassa.";
		case "Cannot delete Asset Base Folder.": return "P\u00E4\u00E4hakemistoa ei voi poistaa";
        }
    }
function loadTxt()
    {
    document.getElementById("btnCloseAndRefresh").value = "Sulje & p\u00E4ivit\u00E4";
    }
function writeTitle()
    {
    document.write("<title>Poista kansio</title>")
    }