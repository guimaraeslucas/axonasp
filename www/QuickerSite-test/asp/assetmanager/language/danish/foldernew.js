function getTxt(s)
    {
    switch(s)
        {
		case "Folder already exists.": return "Mappen eksisterer allerede.";
		case "Folder created.": return "Mappen er oprettet.";
		case "Invalid input.": return "Indtastningen indeholder ugyldige karakterer.";
        }
    }
function loadTxt()
    {
    document.getElementById("txtLang").innerHTML = "Nyt mappe navn";
    document.getElementById("btnCloseAndRefresh").value = "Luk og opdater";
    document.getElementById("btnCreate").value = "Opret";
    }
function writeTitle()
    {
    document.write("<title>Opret mappe</title>")
    }
