package caddy

import (
	"github.com/caddyserver/caddy/v2/caddyconfig/caddyfile"
	"github.com/caddyserver/caddy/v2/caddyconfig/httpcaddyfile"
	"github.com/caddyserver/caddy/v2/modules/caddyhttp"
)

func parseCaddyfile(h httpcaddyfile.Helper) (caddyhttp.MiddlewareHandler, error) {
	var a AxonASP
	err := a.UnmarshalCaddyfile(h.Dispenser)
	return &a, err
}

func (a *AxonASP) UnmarshalCaddyfile(d *caddyfile.Dispenser) error {
	for d.Next() {
		for d.NextBlock(0) {
			switch d.Val() {
			case "site_name":
				if !d.NextArg() {
					return d.ArgErr()
				}
				a.SiteName = d.Val()
			case "config_file":
				if !d.NextArg() {
					return d.ArgErr()
				}
				a.ConfigFile = d.Val()
			case "global_asa_path":
				if !d.NextArg() {
					return d.ArgErr()
				}
				a.GlobalAsaPath = d.Val()
			default:
				return d.Errf("unrecognized subdirective '%s'", d.Val())
			}
		}
	}
	return nil
}
