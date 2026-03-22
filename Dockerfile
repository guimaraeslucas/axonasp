# AxonASP Server - Production Dockerfile
#
# AxonASP Server
# Copyright (C) 2026 G3pix Ltda. All rights reserved.
#
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

# ─── Stage 1: Builder ────────────────────────────────────────────────────────
FROM golang:1.25-alpine AS builder

ARG VERSION=dev
ARG TARGETOS=linux
ARG TARGETARCH=amd64

WORKDIR /build

# Install git (required by some go modules)
RUN apk add --no-cache git

# Cache dependency downloads separately from source
COPY go.mod go.sum ./
RUN go mod download

# Copy source tree
COPY . .

# Build the main server binary (CGO disabled for fully static binary)
RUN CGO_ENABLED=0 GOOS=${TARGETOS} GOARCH=${TARGETARCH} \
    go build -trimpath \
    -ldflags "-s -w -X main.Version=${VERSION}" \
    -o axonasp .

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

# Copy the compiled binary
COPY --from=builder /build/axonasp ./axonasp

# Copy runtime-required assets
COPY --from=builder /build/errorpages ./errorpages
COPY --from=builder /build/www        ./www

# Create required runtime directories
RUN mkdir -p temp/session temp/cache/ast temp/cache/bytecode

# Create a non-root user for security
RUN addgroup -S axonasp && adduser -S axonasp -G axonasp && \
    chown -R axonasp:axonasp /app

USER axonasp

# Default environment (can be overridden at runtime)
ENV SERVER_PORT=4050 \
    WEB_ROOT=./www \
    TIMEZONE=America/Sao_Paulo \
    ASP_CACHE_TYPE=memory \
    DEBUG_ASP=FALSE \
    CLEAN_SESSIONS=TRUE

EXPOSE 4050

HEALTHCHECK --interval=30s --timeout=5s --start-period=10s --retries=3 \
    CMD wget -qO- http://localhost:${SERVER_PORT}/ > /dev/null || exit 1

CMD ["./axonasp"]
