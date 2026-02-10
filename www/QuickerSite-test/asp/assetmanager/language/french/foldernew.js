function getTxt(s)
    {
    switch(s)
        {
		case "Folder already exists.": return "Le r&eacute;pertoire existe d&eacute;j&agrave;.";
		case "Folder created.": return "Le r&eacute;pertoire a &eacute;t&eacute; cr&eacute;&eacute;";
        case "Invalid input.":return "La valeur entr\u00E9e ne convient pas.";//"La valeur entr&eacute;e ne convient pas."
        }
    }   
function loadTxt()
    {
    document.getElementById("txtLang").innerHTML = "Nom du nouveau r\u00E9pertoire";
    document.getElementById("btnCloseAndRefresh").value = "fermer et actualiser";
    document.getElementById("btnCreate").value = "Cr\u00E9er";
    }
function writeTitle()
    {
    document.write("<title>Cr\u00E9ation de r\u00E9pertoire</title>")
    }    
