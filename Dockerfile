# AxonASP Server - Production Dockerfile
#
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
#  must be cited in the documentation or "About" section.
# 
# Contribution Policy:
#  Modifications to the core source code of AxonASP Server must be
#  made available under this same license terms.
#  This code was contributed by @antoniolago (https://github.com/antoniolago)
#

# ─── Stage 1: Builder ────────────────────────────────────────────────────────
FROM golang:1.26.3-alpine AS builder

ARG TARGETOS=linux
ARG TARGETARCH=amd64

WORKDIR /build

# Install git (required to fetch version and for some go modules)
RUN apk add --no-cache git

# Cache dependency downloads separately from source
COPY go.mod go.sum ./
RUN go mod download

# Copy source tree
COPY . .

# Extract version from Git and build all binaries
# Defaulting to 2.0.0.0 if git commands fail (e.g., if not built from a git clone)
RUN PATCH=$(git rev-list --count HEAD 2>/dev/null || echo "0") && \
    REVISION=$(git rev-parse --short HEAD 2>/dev/null || echo "0") && \
    FULL_VERSION="2.0.${PATCH}.${REVISION}" && \
    echo "Building with version: ${FULL_VERSION}" && \
    CGO_ENABLED=0 GOOS=${TARGETOS} GOARCH=${TARGETARCH} go build -trimpath -ldflags "-s -w -X main.Version=${FULL_VERSION}" -o axonasp-http ./server && \
    CGO_ENABLED=0 GOOS=${TARGETOS} GOARCH=${TARGETARCH} go build -trimpath -ldflags "-s -w -X main.Version=${FULL_VERSION}" -o axonasp-fastcgi ./fastcgi && \
    CGO_ENABLED=0 GOOS=${TARGETOS} GOARCH=${TARGETARCH} go build -trimpath -ldflags "-s -w -X main.Version=${FULL_VERSION}" -o axonasp-cli ./cli && \
    CGO_ENABLED=0 GOOS=${TARGETOS} GOARCH=${TARGETARCH} go build -trimpath -ldflags "-s -w -X main.Version=${FULL_VERSION}" -o axonasp-mcp ./mcp && \
    CGO_ENABLED=0 GOOS=${TARGETOS} GOARCH=${TARGETARCH} go build -trimpath -ldflags "-s -w -X main.Version=${FULL_VERSION}" -o axonasp-testsuite ./testsuite

# ─── Stage 2: Runtime ─────────────────────────────────────────────────────────
FROM alpine:3.21

LABEL org.opencontainers.image.title="AxonASP Server"
LABEL org.opencontainers.image.description="G3pix AxonASP - ASP Classic / VBScript web server written in Go"
LABEL org.opencontainers.image.url="https://g3pix.com.br/axonasp"
LABEL org.opencontainers.image.source="https://github.com/guimaraeslucas/axonasp"
LABEL org.opencontainers.image.licenses="MPL-2.0"

# CA certificates for HTTPS outbound calls and tzdata for timezone support
RUN apk add --no-cache ca-certificates tzdata && \
    update-ca-certificates

WORKDIR /app

# Copy all compiled binaries
COPY --from=builder /build/axonasp-http ./axonasp-http
COPY --from=builder /build/axonasp-fastcgi ./axonasp-fastcgi
COPY --from=builder /build/axonasp-cli ./axonasp-cli
COPY --from=builder /build/axonasp-mcp ./axonasp-mcp
COPY --from=builder /build/axonasp-testsuite ./axonasp-testsuite
# Copy runtime-required assets
COPY --from=builder /build/config/ ./config
COPY --from=builder /build/www        ./www
COPY --from=builder /build/mcp        ./mcp
COPY --from=builder /build/LICENSE.txt   ./LICENSE.txt
# For CLI tools, global.asa is required to be in the same directory as the binary
COPY --from=builder /build/global.asa   ./global.asa

# Create required runtime directories
RUN mkdir -p temp/

# Create a non-root user for security
RUN addgroup -S axonasp && adduser -S axonasp -G axonasp && \
    chown -R axonasp:axonasp /app

USER axonasp

# Expose requested ports
# 8801: HTTP Server
# 9000: FastCGI Server
# 8000: MCP Server
EXPOSE 8801 9000 8000

# Healthcheck defaulting to HTTP server port
HEALTHCHECK --interval=30s --timeout=5s --start-period=10s --retries=3 \
    CMD wget -qO- http://localhost:8801/ > /dev/null || exit 1

# Default command runs the HTTP server. 
# Override this command when running the container to start FastCGI or MCP instead.
CMD ["./axonasp-http"]