function getTxt(s)
    {
	switch(s)
        {
		case "Folder already exists.": return "Kansio on jo olemassa.";
		case "Folder created.": return "Kansio luotu.";
        case "Invalid input.":return "Virheellinen sy\u00F6te.";
        }
    }   
function loadTxt()
    {
    document.getElementById("txtLang").innerHTML = "Uuden kansion nimi";
    document.getElementById("btnCloseAndRefresh").value = "Sulje & p\u00E4ivit\u00E4";
    document.getElementById("btnCreate").value = "Luo";
    }
function writeTitle()
    {
    document.write("<title>Luo kansio</title>")
    }