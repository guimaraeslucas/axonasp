function getTxt(s)
    {
    switch(s)
        {
        case "Cannot delete Asset Base Folder.":return "No se puede borrar la carpeta raiz.";
        case "Delete this file ?":return "Borrar el archivo?";
        case "Uploading...":return "Subiendo el archivo...";
        case "File already exists. Do you want to replace it?":return "El archivo ya existe. Quiere reemplazarlo?"

		case "Files": return "Archivos";
		case "del": return "Borrar";
		case "Empty...": return "Vacio...";
        }
    }
function loadTxt()
    {
    var txtLang = document.getElementsByName("txtLang");
    txtLang[0].innerHTML = "Nueva carpeta";
    txtLang[1].innerHTML = "Borrar carpeta";
    txtLang[2].innerHTML = "Subir archivo";

	var optLang = document.getElementsByName("optLang");
    optLang[0].text = "Todos";
    optLang[1].text = "Multimedia";
    optLang[2].text = "Imagenes";
    optLang[3].text = "Flash";

    document.getElementById("btnOk").value = " Aceptar ";
    document.getElementById("btnUpload").value = "Subir";
    }
function writeTitle()
    {
    document.write("<title>Administrador de archivos</title>")
    }