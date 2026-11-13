//go:build !windows

package main

import (
	"context"
	"os"
	"path/filepath"
	"sync"
	"syscall"
	"testing"
	"time"
)

func TestSIGUSR2ReloadsOnlyModifiedPools(t *testing.T) {
	tmpDir := t.TempDir()

	poolAPath := filepath.Join(tmpDir, "site-a.conf")
	poolBPath := filepath.Join(tmpDir, "site-b.conf")

	if err := os.WriteFile(poolAPath, []byte("site_name='a'\nuid=1001\ngid=1001\nsocket='9001'\nconfig_file='/tmp/a.toml'\napp_path='/'\nmemory_limit_mb=64\nmax_restarts=0\n"), 0644); err != nil {
		t.Fatalf("failed to write pool A: %v", err)
	}
	if err := os.WriteFile(poolBPath, []byte("site_name='b'\nuid=1001\ngid=1001\nsocket='9002'\nconfig_file='/tmp/b.toml'\napp_path='/'\nmemory_limit_mb=64\nmax_restarts=0\n"), 0644); err != nil {
		t.Fatalf("failed to write pool B: %v", err)
	}

	originalConfigDir := configDir
	originalActivePools := activePools
	originalLauncher := launchPoolSupervisor

	configDir = tmpDir
	activePools = make(map[string]poolHandle)

	var mu sync.Mutex
	starts := make(map[string]int)
	cancels := make(map[string]int)
	started := make(chan string, 16)

	launchPoolSupervisor = func(ctx context.Context, configPath string, done chan struct{}) {
		mu.Lock()
		starts[configPath]++
		mu.Unlock()
		started <- configPath

		<-ctx.Done()
		mu.Lock()
		cancels[configPath]++
		mu.Unlock()
		close(done)
	}

	waitForStarts := func(expected int) {
		t.Helper()
		deadline := time.After(2 * time.Second)
		for range expected {
			select {
			case <-started:
			case <-deadline:
				t.Fatalf("timed out waiting for %d pool supervisor start(s)", expected)
			}
		}
	}

	t.Cleanup(func() {
		shutdownAllPools()
		configDir = originalConfigDir
		activePools = originalActivePools
		launchPoolSupervisor = originalLauncher
	})

	scanAndLoadConfigs()
	waitForStarts(2)

	mu.Lock()
	if starts[poolAPath] != 1 || starts[poolBPath] != 1 {
		mu.Unlock()
		t.Fatalf("expected both pools to start once, got startsA=%d startsB=%d", starts[poolAPath], starts[poolBPath])
	}
	mu.Unlock()

	time.Sleep(2 * time.Millisecond)
	if err := os.WriteFile(poolAPath, []byte("site_name='a'\nuid=1001\ngid=1001\nsocket='9001'\nconfig_file='/tmp/a.toml'\napp_path='/'\nglobal_asa='/srv/a'\nmemory_limit_mb=64\nmax_restarts=0\n"), 0644); err != nil {
		t.Fatalf("failed to update pool A: %v", err)
	}

	if shouldStop := handleManagerSignal(syscall.SIGUSR2); shouldStop {
		t.Fatal("SIGUSR2 should not request manager shutdown")
	}
	waitForStarts(1)

	mu.Lock()
	startsA := starts[poolAPath]
	startsB := starts[poolBPath]
	cancelsA := cancels[poolAPath]
	cancelsB := cancels[poolBPath]
	mu.Unlock()

	if startsA != 2 {
		t.Fatalf("expected modified pool A to restart once, got starts=%d", startsA)
	}
	if startsB != 1 {
		t.Fatalf("expected unmodified pool B to remain running, got starts=%d", startsB)
	}
	if cancelsA != 1 {
		t.Fatalf("expected modified pool A to receive one cancellation, got %d", cancelsA)
	}
	if cancelsB != 0 {
		t.Fatalf("expected unmodified pool B to not be cancelled, got %d", cancelsB)
	}
}
