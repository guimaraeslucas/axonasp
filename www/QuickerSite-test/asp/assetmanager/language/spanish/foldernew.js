function getTxt(s)
    {
    switch(s)
        {
		case "Folder already exists.": return "La carpeta ya existe.";
		case "Folder created.": return "Carpeta creada.";
        case "Invalid input.":return "Nombre no valido.";
        }
    }
function loadTxt()
    {
    document.getElementById("txtLang").innerHTML = "Nombre nueva carpeta";
    document.getElementById("btnCloseAndRefresh").value = "Cerrar y actualizar";
    document.getElementById("btnCreate").value = "Crear";
    }
function writeTitle()
    {
    document.write("<title>Crear Carpeta</title>")
    }
