#!/usr/bin/env bash
# Build script for Caddy with AxonASP module targeting Linux (amd64)

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" &> /dev/null && pwd)"
cd "$SCRIPT_DIR" || exit 1

PARENT_DIR="$(cd "$SCRIPT_DIR/.." && pwd)"

echo "=== Building Caddy with AxonASP module for Linux (amd64) ==="

if ! command -v xcaddy &> /dev/null; then
    echo "xcaddy not found. Installing xcaddy..."
    go install github.com/caddyserver/xcaddy/cmd/xcaddy@latest
    export PATH="$PATH:$(go env GOPATH)/bin"
fi

export GOOS=linux
export GOARCH=amd64

echo "Compiling Caddy binary: caddy-linux-amd64..."
xcaddy build --output caddy-linux-amd64 \
    --with g3pix.com.br/axonasp/v2/caddy=. \
    --replace "g3pix.com.br/axonasp/v2=$PARENT_DIR" \
    --replace "github.com/google/cel-go=github.com/google/cel-go@v0.20.1"

if [ $? -eq 0 ]; then
    echo "[SUCCESS] Built caddy-linux-amd64 successfully!"
else
    echo "[FAIL] Failed to build Caddy for Linux."
    exit 1
fi
