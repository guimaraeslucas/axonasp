function getTxt(s)
    {
    switch(s)
        {
		case "Folder already exists.": return "Mappen finns redan.";
		case "Folder created.": return "Mappen har skapats.";
        case "Invalid input.":return "Ej till\u00C3\u00A5tna tecken.";
        }
    }   
function loadTxt()
    {
    document.getElementById("txtLang").innerHTML = "Nytt namn";
    document.getElementById("btnCloseAndRefresh").value = "St\u00C3\u00A4ng och uppdatera";
    document.getElementById("btnCreate").value = "Skapa";
    }
function writeTitle()
    {
    document.write("<title>Ny mapp</title>")
    }