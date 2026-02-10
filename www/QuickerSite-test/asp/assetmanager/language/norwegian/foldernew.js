function getTxt(s)
    {
    switch(s)
        {
		case "Folder already exists.": return "Mappen eksisterer allerede.";
		case "Folder created.": return "Mappen er opprettet.";
        case "Invalid input.":return "Ugyldig inntasting.";
        }
    }   
function loadTxt()
    {
    document.getElementById("txtLang").innerHTML = "Ny mappe navn";
    document.getElementById("btnCloseAndRefresh").value = "Lukk og oppdater";
    document.getElementById("btnCreate").value = "Opprett";
    }
function writeTitle()
    {
    document.write("<title>Opprett mappe</title>")
    }
