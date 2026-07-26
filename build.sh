#!/usr/bin/env bash

#                 AxonASP Build Script
#
# AxonASP Server
# Copyright (C) 2026 G3pix Ltda. All rights reserved.
#
# Developed by Lucas Guimarães - G3pix Ltda
# Contact: https://g3pix.com.br
# Project URL: https://g3pix.com.br/axonasp
#
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.
#
# Attribution Notice:
# If this software is used in other projects, the name "AxonASP Server"
# must be cited in the documentation or "About" section.
#
# Contribution Policy:
# Modifications to the core source code of AxonASP Server must be
# made available under this same license terms.
#

# --- Defaults ---
PLATFORM="linux"
ARCHITECTURE="amd64"
CLEAN=0
TEST=0
TAGS=""

# --- Argument Parsing ---
while [[ "$#" -gt 0 ]]; do
    case $1 in
        --platform|-p) PLATFORM="$2"; shift ;;
        --arch|-a) ARCHITECTURE="$2"; shift ;;
        --clean|-c) CLEAN=1 ;;
        --test|-t) TEST=1 ;;
        --tags|-g) TAGS="$2"; shift ;;
        *) echo -e "\033[0;31mUnknown parameter passed: $1\033[0m"; exit 1 ;;
    esac
    shift
done

# Normalize Go build tags so users can pass comma/semicolon/space-separated values.
NORMALIZED_TAGS=$(echo "$TAGS" | tr ',;' '  ' | xargs)

# --- Validate Parameters ---
if [[ ! "$PLATFORM" =~ ^(windows|linux|darwin|wasm|all)$ ]]; then
    echo "Invalid platform: $PLATFORM. Allowed: windows, linux, darwin, wasm, all."
    exit 1
fi
if [[ ! "$ARCHITECTURE" =~ ^(amd64|arm64|386)$ ]]; then
    echo "Invalid architecture: $ARCHITECTURE. Allowed: amd64, arm64, 386."
    exit 1
fi

# --- AUTOMATIC VERSION CONFIGURATION ---
MAJOR="2"
MINOR="2"
PATCH="0"
REVISION="0"

if git rev-parse --is-inside-work-tree >/dev/null 2>&1; then
    GIT_TAG=$(git describe --tags --abbrev=0 2>/dev/null)
    
    REGEX="^v?([0-9]+)\.([0-9]+)\.([0-9]+)$"
    
    if [[ $GIT_TAG =~ $REGEX ]]; then
        MAJOR="${BASH_REMATCH[1]}"
        MINOR="${BASH_REMATCH[2]}"
        PATCH="${BASH_REMATCH[3]}"
    else
        PATCH=$(git rev-list --count HEAD | xargs)
    fi

    REVISION=$(git rev-parse --short HEAD | xargs)
else
    echo -e "\033[1;33mGit not found or not a valid repository. Using default versioning.\033[0m"
fi

FULL_VERSION="$MAJOR.$MINOR.$PATCH.$REVISION"

# --- Color output functions ---
GREEN='\033[0;32m'
CYAN='\033[0;36m'
RED='\033[0;31m'
YELLOW='\033[1;33m'
MAGENTA='\033[0;35m'
WHITE='\033[1;37m'
DARKGRAY='\033[1;30m'
NC='\033[0m' # No Color

write_success() { echo -e "${GREEN}$1${NC}"; }
write_info() { echo -e "${CYAN}$1${NC}"; }
write_err() { echo -e "${RED}$1${NC}"; }
write_warn() { echo -e "${YELLOW}$1${NC}"; }

# Script header
echo ""
echo -e "${MAGENTA}=======================================================${NC}"
echo -e " ${WHITE} G3Pix ❖ AxonASP Build Script${NC}"
echo -e " ${CYAN} Version: $FULL_VERSION${NC}"
if [ -n "$NORMALIZED_TAGS" ]; then
    echo -e " ${YELLOW} Build Tags: $NORMALIZED_TAGS${NC}"
fi
echo -e "${MAGENTA}=======================================================${NC}"
echo ""

# Set Working Directory to script location
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" &> /dev/null && pwd)"
cd "$SCRIPT_DIR" || exit 1

# Clean previous builds
if [ "$CLEAN" -eq 1 ]; then
    write_info "Cleaning previous builds..."
    rm -f axonasp-http.exe axonasp-fastcgi.exe axonasp-cli.exe axonasp-testsuite.exe axonasp-mcp.exe axonasp-service.exe axonasp-admin.exe axonhta.exe axonasp-http axonasp-fastcgi axonasp-cli axonasp-testsuite axonasp-mcp axonasp-service axonasp-admin axonhta
    rm -rf build
    write_success "Cleaned."
    echo ""
fi

# Targets
TARGET_LABELS=("HTTP Server" "FastCGI Server" "CLI" "Test Suite" "MCP" "Service Wrapper" "Admin Tool" "HTA Desktop")
TARGET_OUTPUTS=("axonasp-http" "axonasp-fastcgi" "axonasp-cli" "axonasp-testsuite" "axonasp-mcp" "axonasp-service" "axonasp-admin" "axonhta")
TARGET_SOURCES=("./server" "./fastcgi" "./cli" "./testsuite" "./mcp" "./service" "./admin" "./axonhta")

BUILD_SUCCESS=true

build_binary() {
    local target_os="$1"
    local target_arch="$2"
    local output_name="$3"
    local source_path="$4"
    local label="$5"

    export GOOS="$target_os"
    export GOARCH="$target_arch"

    local extension=""
    if [ "$target_os" == "windows" ]; then
        extension=".exe"
    fi

    local output_file="${output_name}${extension}"
    local ldflags="-X main.Version=$FULL_VERSION"
    # AxonHTA is a desktop app: hide the console window on Windows.
    if [ "$output_name" = "axonhta" ] && [ "$target_os" = "windows" ]; then
        ldflags="$ldflags -H windowsgui"
    fi
    local build_args=(build -trimpath -ldflags "$ldflags" -o "$output_file")
    if [ -n "$NORMALIZED_TAGS" ]; then
        build_args+=(-tags "$NORMALIZED_TAGS")
    fi
    build_args+=("$source_path")

    write_info "Building $label ($target_os/$target_arch) -> $output_file ..."

    # Build and capture output
    local build_output
    build_output=$(go "${build_args[@]}" 2>&1)
    local exit_code=$?

    if [ $exit_code -eq 0 ] && [ -f "$output_file" ]; then
        local bytes
        bytes=$(wc -c < "$output_file")
        local size_mb
        size_mb=$(awk "BEGIN {printf \"%.2f\", $bytes / 1048576}")
        write_success "  [OK] $output_file ($size_mb MB)"
        return 0
    else
        write_err "  [FAIL] $label"
        if [ -n "$build_output" ]; then echo "$build_output"; fi
        return 1
    fi
}

# --- Fix, format and generate before any build pass ---
write_info "Fixing source..."
go fix ./... > /dev/null 2>&1

write_info "Formatting source..."
gofmt -w . > /dev/null 2>&1

write_info "Running go generate..."
go generate ./... > /dev/null 2>&1
echo ""

run_platform() {
    local os="$1"
    local arch="$2"
    local out_dir=""

    if [ "$os" != "windows" ]; then
        out_dir="build/$os-$arch/"
        mkdir -p "$out_dir"
    fi

    echo -e "${DARKGRAY}-------------------------------------------------------${NC}"
    echo -e " ${YELLOW}Building for $os/$arch${NC}"
    echo -e "${DARKGRAY}-------------------------------------------------------${NC}"

    for i in "${!TARGET_LABELS[@]}"; do
        local out="${out_dir}${TARGET_OUTPUTS[$i]}"
        build_binary "$os" "$arch" "$out" "${TARGET_SOURCES[$i]}" "${TARGET_LABELS[$i]}"
        if [ $? -ne 0 ]; then
            BUILD_SUCCESS=false
        fi
    done

    if [ "$os" != "windows" ]; then
        local fpm_out="${out_dir}axonasp-fpm"
        build_binary "$os" "$arch" "$fpm_out" "./fpm" "FPM Manager"
        if [ $? -ne 0 ]; then
            BUILD_SUCCESS=false
        fi
    fi
    echo ""
}

# Execute Platform Builds
if [[ "$PLATFORM" == "windows" || "$PLATFORM" == "all" ]]; then run_platform "windows" "$ARCHITECTURE"; fi
if [[ "$PLATFORM" == "linux" || "$PLATFORM" == "all" ]]; then run_platform "linux" "$ARCHITECTURE"; fi
if [[ "$PLATFORM" == "darwin" || "$PLATFORM" == "all" ]]; then run_platform "darwin" "$ARCHITECTURE"; fi

if [[ "$PLATFORM" == "wasm" || "$PLATFORM" == "all" ]]; then
    echo -e "${DARKGRAY}-------------------------------------------------------${NC}"
    echo -e " ${YELLOW}Building for js/wasm${NC}"
    echo -e "${DARKGRAY}-------------------------------------------------------${NC}"

    out_dir="./wasm/"
    build_binary "js" "wasm" "${out_dir}axonasp.wasm" "./wasm/main.go" "WASM Web"
    if [ $? -ne 0 ]; then BUILD_SUCCESS=false; fi

    GOROOT=$(go env GOROOT)
    cp "$GOROOT/lib/wasm/wasm_exec.js" "$out_dir"
    write_success "  [OK] Copied wasm_exec.js"
    echo ""
fi

# Reset environment variables to native host
unset GOOS
unset GOARCH

# --- Tests ---
if [ "$TEST" -eq 1 ]; then
    echo -e "${DARKGRAY}-------------------------------------------------------${NC}"
    echo -e " ${YELLOW}Running Tests${NC}"
    echo -e "${DARKGRAY}-------------------------------------------------------${NC}"
    echo ""

    write_info "Running go test ./..."
    test_output=$(go test ./... 2>&1)
    local test_exit=$?

    if [ $test_exit -eq 0 ]; then
        write_success "[OK] All tests passed"
    else
        write_err "[FAIL] Some tests failed"
        echo "$test_output"
        BUILD_SUCCESS=false
    fi
    echo ""
fi

# --- Summary ---
echo -e "${MAGENTA}=======================================================${NC}"

if [ "$BUILD_SUCCESS" = true ]; then
    write_success "  BUILD SUCCESSFUL  (v$FULL_VERSION)"
    echo ""
    echo -e " ${WHITE} Executables:${NC}"

    # List root executables
    for file in axonasp-http axonasp-fastcgi axonasp-cli axonasp-testsuite axonasp-mcp axonasp-service axonasp-admin axonhta; do
        if [ -f "$file" ]; then echo -e "    - ${CYAN}$file${NC}"; fi
    done

    # List build directory executables
    if [ -d "build" ]; then
        find build -type f | while read -r file; do
            echo -e "    - ${CYAN}$file${NC}"
        done
    fi

    echo ""
    echo -e " ${WHITE} Quick Start:${NC}"
    echo -e "    ${DARKGRAY}HTTP Server : ./axonasp-http${NC}"
    echo -e "    ${DARKGRAY}FastCGI     : ./axonasp-fastcgi${NC}"
    echo -e "    ${DARKGRAY}CLI         : ./axonasp-cli${NC}"
    echo -e "    ${DARKGRAY}Test Suite  : ./axonasp-testsuite ./www/tests${NC}"
    echo -e "    ${DARKGRAY}MCP         : ./axonasp-mcp${NC}"
    echo -e "    ${DARKGRAY}Service     : ./axonasp-service install|start|stop|uninstall${NC}"
    echo -e "    ${DARKGRAY}Admin Tool  : ./axonasp-admin${NC}"
    echo -e "    ${DARKGRAY}HTA Desktop : ./axonhta --app ./axonhta/HTAtest/axonhta-todo-app${NC}"
    
    echo -e "${MAGENTA}=======================================================${NC}"
    echo ""
    exit 0
else
    write_err "  BUILD FAILED!"
    echo ""
    echo -e "  ${YELLOW}Check the error messages above for details.${NC}"
    echo -e "${MAGENTA}=======================================================${NC}"
    echo ""
    exit 1
fi