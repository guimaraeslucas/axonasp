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
FROM golang:1.26-alpine AS builder

ARG TARGETOS=linux
ARG TARGETARCH=amd64
# Allows CI to pass the version string directly via --build-arg VERSION=...
ARG VERSION 

WORKDIR /build

# Install git (required to fetch version and for some go modules)
RUN apk add --no-cache git

# Cache dependency downloads separately from source
COPY go.mod go.sum ./
RUN go mod download

# Copy source tree
COPY . .

# Extract version from Git if not provided via ARG, and build all binaries
RUN if [ -z "$VERSION" ]; then \
    PATCH=$(git rev-list --count HEAD 2>/dev/null || echo "0"); \
    REVISION=$(git rev-parse --short HEAD 2>/dev/null || echo "0"); \
    VERSION="2.3.${PATCH}.${REVISION}"; \
    fi && \
    echo "Building with version: ${VERSION}" && \
    CGO_ENABLED=0 GOOS=${TARGETOS} GOARCH=${TARGETARCH} go build -trimpath -ldflags "-s -w -X main.Version=${VERSION}" -o axonasp-http ./server && \
    CGO_ENABLED=0 GOOS=${TARGETOS} GOARCH=${TARGETARCH} go build -trimpath -ldflags "-s -w -X main.Version=${VERSION}" -o axonasp-fastcgi ./fastcgi && \
    CGO_ENABLED=0 GOOS=${TARGETOS} GOARCH=${TARGETARCH} go build -trimpath -ldflags "-s -w -X main.Version=${VERSION}" -o axonasp-cli ./cli && \
    CGO_ENABLED=0 GOOS=${TARGETOS} GOARCH=${TARGETARCH} go build -trimpath -ldflags "-s -w -X main.Version=${VERSION}" -o axonasp-mcp ./mcp && \
    CGO_ENABLED=0 GOOS=${TARGETOS} GOARCH=${TARGETARCH} go build -trimpath -ldflags "-s -w -X main.Version=${VERSION}" -o axonasp-admin ./admin && \
    CGO_ENABLED=0 GOOS=${TARGETOS} GOARCH=${TARGETARCH} go build -trimpath -ldflags "-s -w -X main.Version=${VERSION}" -o axonasp-fpm ./fpm && \
    CGO_ENABLED=0 GOOS=${TARGETOS} GOARCH=${TARGETARCH} go build -trimpath -ldflags "-s -w -X main.Version=${VERSION}" -o axonasp-testsuite ./testsuite


# ─── Stage 2: Runtime ─────────────────────────────────────────────────────────
FROM alpine:3.21

LABEL org.opencontainers.image.title="AxonASP Server"
LABEL org.opencontainers.image.description="AxonASP Server - ASP Classic web server with Proxy, FastCGI, CLI, TestSuite and MCP support. VBScript and JavaScript support. Small, fast and secure alternative to IIS for hosting ASP applications on modern platforms."
LABEL org.opencontainers.image.url="https://g3pix.com.br/axonasp"
LABEL org.opencontainers.image.source="https://github.com/guimaraeslucas/axonasp"
LABEL org.opencontainers.image.licenses="MPL-2.0"

# Adicionamos o su-exec para fazer o drop de privilégios dinamicamente
RUN apk add --no-cache ca-certificates tzdata su-exec && \
    update-ca-certificates && \
    addgroup -S axonasp && adduser -S axonasp -G axonasp

WORKDIR /opt/axonasp

# Copy binaries and assets directly with the correct ownership
COPY --from=builder /build/axonasp-http ./axonasp-http
COPY --from=builder /build/axonasp-fastcgi ./axonasp-fastcgi
COPY --from=builder /build/axonasp-cli ./axonasp-cli
COPY --from=builder /build/axonasp-mcp ./axonasp-mcp
COPY --from=builder /build/axonasp-fpm ./axonasp-fpm
COPY --from=builder /build/axonasp-testsuite ./axonasp-testsuite
COPY --from=builder /build/axonasp-admin ./axonasp-admin


COPY --from=builder /build/config/ ./config/
COPY --from=builder /build/fpm/fpm.d/ ./fpm/default_fpm.d/
COPY --from=builder /build/mcp/ ./mcp/
COPY --from=builder /build/LICENSE.txt ./LICENSE.txt
COPY --from=builder /build/global.asa ./global.asa
COPY --from=builder  /build/www/ ./default_www/

# Create required runtime directories
RUN mkdir -p temp/ www/

RUN printf '#!/bin/sh\n\
    set -e\n\
    if [ -z "$(ls -A /opt/axonasp/www 2>/dev/null)" ]; then\n\
    echo "--> Initializing /opt/axonasp/www with the default AxonASP content..."\n\
    cp -a /opt/axonasp/default_www/. /opt/axonasp/www/\n\
    fi\n\
    if [ -z "$(ls -A /opt/axonasp/fpm/fpm.d/ 2>/dev/null)" ]; then\n\
    echo "--> Initializing /opt/axonasp/fpm/fpm.d/ with the default AxonASP fpm configuration..."\n\
    cp -a /opt/axonasp/fpm/default_fpm.d/. /opt/axonasp/fpm/fpm.d/\n\
    fi\n\
    # 2. Ensure the directories belong to the axonasp user\n\
    chown -R axonasp:axonasp /opt/axonasp/www\n\
    chown -R axonasp:axonasp /opt/axonasp/temp\n\
    chown -R axonasp:axonasp /opt/axonasp/fpm/fpm.d\n\
    # 3. Switch from root to axonasp user and execute the main binary\n\
    exec su-exec axonasp "$@"\n' > /entrypoint.sh && chmod +x /entrypoint.sh

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
ENTRYPOINT ["/entrypoint.sh"]
CMD ["./axonasp-http"]
