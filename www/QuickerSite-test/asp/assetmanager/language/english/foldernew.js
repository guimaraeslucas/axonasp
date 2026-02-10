function getTxt(s)
	{
	switch(s)
		{
		case "Folder already exists.": return "Folder already exists.";
		case "Folder created.": return "Folder created.";
		case "Invalid input.": return "Invalid input.";
		}
	}	
function loadTxt()
	{
    document.getElementById("txtLang").innerHTML = "New Folder Name";
    document.getElementById("btnCloseAndRefresh").value = "close & refresh";
    document.getElementById("btnCreate").value = "create";
	}
function writeTitle()
	{
	document.write("<title>Create Folder</title>")
	}
