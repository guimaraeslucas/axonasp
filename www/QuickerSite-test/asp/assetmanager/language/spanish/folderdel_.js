function getTxt(s)
	{
    switch(s)
        {
		case "Folder deleted.": return "Carpeta borrada.";
		case "Folder does not exist.": return "La carpeta no existe.";
		case "Cannot delete Asset Base Folder.": return "No se puede borrar la carpeta.";
        }
    }
function loadTxt()
    {
    document.getElementById("btnCloseAndRefresh").value = "Cerrar y actualizar";
    }
function writeTitle()
    {
    document.write("<title>Borrar Carpeta</title>")
    }