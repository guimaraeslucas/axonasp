package main

import (
	caddycmd "github.com/caddyserver/caddy/v2/cmd"

	// plug in Caddy modules here
	_ "github.com/caddyserver/caddy/v2/modules/standard"
	_ "g3pix.com.br/axonasp/caddy"
)

func main() {
	caddycmd.Main()
}