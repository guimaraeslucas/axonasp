function getTxt(s)
	{
	switch(s)
		{
		case "Folder already exists.": return "Cartella già esistente.";
		case "Folder created.": return "Cartella Creata.";
		case "Invalid input.": return "Input invalido.";
		}
	}	
function loadTxt()
	{
    document.getElementById("txtLang").innerHTML = "Nuovo nome cartella";
    document.getElementById("btnCloseAndRefresh").value = "chiudi & rinfresca";
    document.getElementById("btnCreate").value = "crea";
	}
function writeTitle()
	{
	document.write("<title>Crea Cartella</title>")
	}
