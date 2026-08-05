package axonvm

import (
	"fmt"
	"os"
	"path/filepath"
	"regexp"
	"runtime"
	"strings"
	"testing"

	"github.com/shirou/gopsutil/v3/disk"
)

// TestFSODiskUsageMatchesGopsutil verifies the FSO disk usage adapter preserves
// the underlying filesystem totals exposed by gopsutil.
func TestFSODiskUsageMatchesGopsutil(t *testing.T) {
	t.Parallel()

	rootPath := string(os.PathSeparator)
	if runtime.GOOS == "windows" {
		cwd, err := os.Getwd()
		if err != nil {
			t.Fatalf("get cwd: %v", err)
		}
		volume := filepath.VolumeName(cwd)
		if volume == "" {
			t.Skip("no drive volume available for disk usage test")
		}
		rootPath = volume + string(os.PathSeparator)
	}

	expected, err := disk.Usage(rootPath)
	if err != nil {
		t.Fatalf("disk usage for %q: %v", rootPath, err)
	}

	usage := (&VM{}).fsoDiskUsage(rootPath)
	if usage == nil {
		t.Fatalf("expected disk usage state for %q", rootPath)
	}

	if got := usage.Size(); got != expected.Total {
		t.Fatalf("total mismatch for %q: got %d want %d", rootPath, got, expected.Total)
	}
	if got := usage.Free(); got != expected.Free {
		t.Fatalf("free mismatch for %q: got %d want %d", rootPath, got, expected.Free)
	}
	if got := usage.Available(); got != expected.Free {
		t.Fatalf("available mismatch for %q: got %d want %d", rootPath, got, expected.Free)
	}
}

func TestFSOGetTempNameFormat(t *testing.T) {
	re := regexp.MustCompile(`^rad[0-9a-fA-F]{8}\.tmp$`)

	vbsSource := `<%@ Language="VBScript" %>` +
		`<%` +
		`Set fso = Server.CreateObject("Scripting.FileSystemObject")` +
		`Response.Write fso.GetTempName` +
		`%>`
	vbsOut := runASPSourceForTest(t, vbsSource)
	if len(vbsOut) != 15 {
		t.Fatalf("expected VBScript GetTempName length 15, got %d (%q)", len(vbsOut), vbsOut)
	}
	if !re.MatchString(vbsOut) {
		t.Fatalf("expected VBScript GetTempName to match ^rad[0-9a-fA-F]{8}\\.tmp$, got %q", vbsOut)
	}

	jsSource := `<%@ Language="JScript" %>` +
		`<%` +
		`var fso = Server.CreateObject("Scripting.FileSystemObject");` +
		`Response.Write(fso.GetTempName());` +
		`%>`
	jsOut := runASPSourceForTest(t, jsSource)
	if len(jsOut) != 15 {
		t.Fatalf("expected JScript GetTempName length 15, got %d (%q)", len(jsOut), jsOut)
	}
	if !re.MatchString(jsOut) {
		t.Fatalf("expected JScript GetTempName to match ^rad[0-9a-fA-F]{8}\\.tmp$, got %q", jsOut)
	}
}

func TestFSOGetSpecialFolderFolderObject(t *testing.T) {
	vbsSource := `<%@ Language="VBScript" %>` +
		`<%` +
		`Set fso = Server.CreateObject("Scripting.FileSystemObject")` +
		`Response.Write fso.GetSpecialFolder(0).Path & "|"` +
		`Response.Write fso.GetSpecialFolder(1).Path & "|"` +
		`Response.Write fso.GetSpecialFolder(2).Path` +
		`%>`
	vbsOut := runASPSourceForTest(t, vbsSource)
	parts := strings.Split(vbsOut, "|")
	if len(parts) != 3 {
		t.Fatalf("expected 3 parts from GetSpecialFolder test, got %v (%q)", parts, vbsOut)
	}

	if parts[2] == "" || parts[2] != os.TempDir() {
		t.Fatalf("expected TempFolder %q, got %q", os.TempDir(), parts[2])
	}

	if runtime.GOOS == "windows" {
		if parts[0] == "" || parts[1] == "" {
			t.Fatalf("expected non-empty Windows/System folders on Windows, got %q | %q", parts[0], parts[1])
		}
	} else {
		if parts[0] != "/usr/bin" {
			t.Fatalf("expected /usr/bin for folder 0 on Unix, got %q", parts[0])
		}
		if parts[1] != "/usr/sbin" {
			t.Fatalf("expected /usr/sbin for folder 1 on Unix, got %q", parts[1])
		}
	}
}

func TestFSOFolderAndFileMoveStateMutation(t *testing.T) {
	tempDir := t.TempDir()
	origFolder := filepath.Join(tempDir, "orig_folder")
	movedFolder := filepath.Join(tempDir, "moved_folder")
	if err := os.MkdirAll(origFolder, 0755); err != nil {
		t.Fatalf("mkdir orig: %v", err)
	}

	jsSource := fmt.Sprintf(`<%%@ Language="JScript" %%><%%
var fso = Server.CreateObject("Scripting.FileSystemObject");
var fld = fso.GetFolder(%q);
fld.Move(%q);
Response.Write(fld.Path + "|");
Response.Write(fld.Name + "|");
var tf = fld.CreateTextFile("readme.txt", true);
tf.Close();
Response.Write(fld.Files.Count + "|");
var fl = fso.GetFile(fld.Path + "/readme.txt");
var movedFile = fld.Path + "/renamed.txt";
fl.Move(movedFile);
Response.Write(fl.Path + "|");
Response.Write(fl.Name);
%%>`, origFolder, movedFolder)

	out := runASPSourceForTest(t, jsSource)
	parts := strings.Split(out, "|")
	if len(parts) != 5 {
		t.Fatalf("expected 5 parts output, got %v (raw: %q)", parts, out)
	}

	if filepath.Clean(parts[0]) != filepath.Clean(movedFolder) {
		t.Fatalf("expected folder path %q, got %q", movedFolder, parts[0])
	}
	if parts[1] != "moved_folder" {
		t.Fatalf("expected folder name %q, got %q", "moved_folder", parts[1])
	}
	if parts[2] != "1" {
		t.Fatalf("expected 1 file in moved folder, got %q", parts[2])
	}
	expectedFilePath := filepath.Join(movedFolder, "renamed.txt")
	if filepath.Clean(parts[3]) != filepath.Clean(expectedFilePath) {
		t.Fatalf("expected file path %q, got %q", expectedFilePath, parts[3])
	}
	if parts[4] != "renamed.txt" {
		t.Fatalf("expected file name %q, got %q", "renamed.txt", parts[4])
	}
}
