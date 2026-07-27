<%@ Language="JScript" %>
<%
var G3PIX_APP_NAME = "G3Pix CMS";
// Change this value to deploy G3Pix under a different virtual directory.
var G3PIX_ROOT_PATH = "/g3pix-cms";
var G3PIX_UPLOAD_DIR = G3PIX_ROOT_PATH + "/uploads/images";
var G3PIX_REMOTE_AXON_CSS = "https://g3pix.com.br/axonasp/css/axonasp.css";

function GetRootFolderName() {
    var normalized = String(G3PIX_ROOT_PATH);
    normalized = normalized.replace(/\\/g, "/");
    normalized = normalized.replace(/\/+$/g, "");
    if (normalized == "") {
        return "g3pix";
    }
    var lastSlash = normalized.lastIndexOf("/");
    if (lastSlash >= 0 && lastSlash < (normalized.length - 1)) {
        return normalized.substring(lastSlash + 1);
    }
    return normalized;
}

function ResolveExistingFolder(candidates) {
    var fso = Server.CreateObject("Scripting.FileSystemObject");
    var found = "";
    var i;
    for (i = 0; i < candidates.length; i++) {
        if (fso.FolderExists(candidates[i])) {
            found = candidates[i];
            break;
        }
    }
    fso = null;
    return found;
}

function ResolveDataDirectoryFS() {
    var rootFolderName = GetRootFolderName();
    var candidates = [
        Server.MapPath(G3PIX_ROOT_PATH + "/data"),
        Server.MapPath("/../" + rootFolderName + "/data"),
        Server.MapPath("./data"),
        Server.MapPath("../data")
    ];
    var resolved = ResolveExistingFolder(candidates);
    if (resolved == "") {
        resolved = candidates[0];
    }
    return resolved;
}

function ResolveUploadDirectoryFS() {
    return Server.MapPath(G3PIX_UPLOAD_DIR);
}

var G3PIX_DATA_DIR_FS = ResolveDataDirectoryFS();
var G3PIX_DB_FILE = G3PIX_DATA_DIR_FS + "/g3pix.sqlite";
var G3PIX_UPLOAD_DIR_FS = ResolveUploadDirectoryFS();

function GetAxonCssHref() {
    var fso = Server.CreateObject("Scripting.FileSystemObject");
    var localCssPath = Server.MapPath("/css/axonasp.css");
    var cssHref = "/css/axonasp.css";

    if (!fso.FileExists(localCssPath)) {
        cssHref = G3PIX_REMOTE_AXON_CSS;
    }

    fso = null;
    return cssHref;
}

function EnsureDataFolder() {
    var fso = Server.CreateObject("Scripting.FileSystemObject");
    if (!fso.FolderExists(G3PIX_DATA_DIR_FS)) {
        fso.CreateFolder(G3PIX_DATA_DIR_FS);
    }
    fso = null;
}

function EnsureUploadFolder() {
    var fso = Server.CreateObject("Scripting.FileSystemObject");
    if (!fso.FolderExists(G3PIX_UPLOAD_DIR_FS)) {
        fso.CreateFolder(G3PIX_UPLOAD_DIR_FS);
    }
    fso = null;
}
%>