//go:build lib_g3axonlive_disabled

/*
 * AxonASP Server
 * Copyright (C) 2026 G3pix Ltda. All rights reserved.
 *
 * Developed by Lucas Guimaraes - G3pix Ltda
 * Contact: https://g3pix.com.br
 * Project URL: https://g3pix.com.br/axonasp
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *
 * Attribution Notice:
 * If this software is used in other projects, the name "AxonASP Server"
 * must be cited in the documentation or "About" section.
 *
 * Contribution Policy:
 * Modifications to the core source code of AxonASP Server must be
 * made available under this same license terms.
 */

// Package main provides the no-op G3AxonLive endpoint stub used when the
// lib_g3axonlive_disabled build tag is set. No routes are registered, no
// goroutines are spawned, and no memory is allocated for the library.
package main

import "net/http"

// RegisterG3AxonLiveEndpoint is a no-op stub when the lib_g3axonlive_disabled build tag is set.
// The /g3al/ endpoint is never registered, ensuring the library consumes zero resources.
func RegisterG3AxonLiveEndpoint(_ *http.ServeMux) {}
