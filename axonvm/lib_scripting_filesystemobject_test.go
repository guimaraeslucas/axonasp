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

	// GetTempName is a method, so the VBScript call site must use parentheses. The
	// parenthesis-free member form resolves as a property read and returns Empty.
	vbsSource := `<%@ Language="VBScript" %>` +
		`<%` +
		`Set fso = Server.CreateObject("Scripting.FileSystemObject")` +
		`Response.Write fso.GetTempName()` +
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

func TestFSOFileExistsAcceptsWindowsSeparatorsOnEveryHost(t *testing.T) {
	vm := NewVM(nil, nil, 5)
	host := NewMockHost()
	rootDir := t.TempDir()
	host.Server().SetRootDir(rootDir)
	host.Server().SetRequestPath("/tests/test_server.asp")
	vm.SetHost(host)

	filePath := filepath.Join(rootDir, "images", "favicon.ico")
	if err := os.MkdirAll(filepath.Dir(filePath), 0755); err != nil {
		t.Fatalf("create image directory: %v", err)
	}
	if err := os.WriteFile(filePath, []byte("icon"), 0644); err != nil {
		t.Fatalf("create image file: %v", err)
	}

	fso := vm.dispatchNativeCall(nativeObjectServer, "CreateObject", []Value{NewString("Scripting.FileSystemObject")})
	if fso.Type != VTNativeObject {
		t.Fatalf("expected VTNativeObject for FSO, got %#v", fso)
	}

	windowsStylePath := host.Server().MapPath("/images") + "\\favicon.ico"
	exists := vm.dispatchNativeCall(fso.Num, "FileExists", []Value{NewString(windowsStylePath)})
	if exists.Type != VTBool || exists.Num != 1 {
		t.Fatalf("expected FileExists True for %q, got %#v", windowsStylePath, exists)
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
	// The FSO sandbox only resolves paths inside the web root, so the temporary
	// sandbox doubles as the web root and the script addresses folders below it.
	rootDir := t.TempDir()
	origFolder := filepath.Join(rootDir, "orig_folder")
	movedFolder := filepath.Join(rootDir, "moved_folder")
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

	host := NewMockHost()
	host.Server().SetRootDir(rootDir)
	host.Server().SetRequestPath("/tests/test_server.asp")
	out := runASPSourceForTestWithHost(t, jsSource, host)
	parts := strings.Split(out, "|")
	if len(parts) != 5 {
		t.Fatalf("expected 5 parts output, got %v (raw: %q)", parts, out)
	}

	if !strings.EqualFold(filepath.Clean(parts[0]), filepath.Clean(movedFolder)) {
		t.Fatalf("expected folder path %q, got %q", movedFolder, parts[0])
	}
	if parts[1] != "moved_folder" {
		t.Fatalf("expected folder name %q, got %q", "moved_folder", parts[1])
	}
	if parts[2] != "1" {
		t.Fatalf("expected 1 file in moved folder, got %q", parts[2])
	}
	expectedFilePath := filepath.Join(movedFolder, "renamed.txt")
	if !strings.EqualFold(filepath.Clean(parts[3]), filepath.Clean(expectedFilePath)) {
		t.Fatalf("expected file path %q, got %q", expectedFilePath, parts[3])
	}
	if parts[4] != "renamed.txt" {
		t.Fatalf("expected file name %q, got %q", "renamed.txt", parts[4])
	}
}

// TestFSOResolvePathVirtualAbsoluteAndSandbox verifies virtual root paths like "/sandbox/x",
// already-mapped physical paths inside root, and traversal rejection across platforms.
func TestFSOResolvePathVirtualAbsoluteAndSandbox(t *testing.T) {
	rootDir := t.TempDir()
	host := NewMockHost()
	host.Server().SetRootDir(rootDir)
	host.Server().SetRequestPath("/sub/page.asp")

	vm := NewVM(nil, nil, 0)
	vm.SetHost(host)

	// Virtual absolute path must resolve under web root on both Windows and Unix.
	resolvedVirtual, ok := vm.fsoResolvePath("/sandbox/x")
	if !ok {
		t.Fatalf("expected virtual absolute path /sandbox/x to resolve inside root %q", rootDir)
	}
	wantVirtual := filepath.Join(rootDir, "sandbox", "x")
	if filepath.Clean(resolvedVirtual) != filepath.Clean(wantVirtual) {
		t.Fatalf("virtual path mismatch: got %q, want %q", resolvedVirtual, wantVirtual)
	}

	// Already-mapped physical path under web root must not be prefixed twice.
	resolvedMapped, ok := vm.fsoResolvePath(wantVirtual)
	if !ok {
		t.Fatalf("expected physical path %q inside root to resolve", wantVirtual)
	}
	if filepath.Clean(resolvedMapped) != filepath.Clean(wantVirtual) {
		t.Fatalf("physical path inside root mismatch: got %q, want %q", resolvedMapped, wantVirtual)
	}

	// Traversal above web root must be refused for both virtual and physical paths.
	if escaped, ok := vm.fsoResolvePath("/../outside.txt"); ok {
		t.Fatalf("expected /../outside.txt to be refused, got %q", escaped)
	}
	if escaped, ok := vm.fsoResolvePath(filepath.Join(rootDir, "..", "outside.txt")); ok {
		t.Fatalf("expected physical root traversal to be refused, got %q", escaped)
	}
}
