# G3AxonLive Configuration

G3AxonLive's behavior can be customized through the main `axonasp.toml` configuration file. These settings allow you to fine-tune performance and resource usage.

All G3AxonLive settings are located under the `[g3axonlive]` section of the `axonasp.toml` file.

```toml
# --- G3AxonLive Settings ---
[g3axonlive]

# Determines if the G3AxonLive library and its /g3al/ endpoint are enabled.
# If false, Server.CreateObject("G3AXONLIVE") will fail and the endpoint will
# return a 404 Not Found. This requires a server restart to take effect.
# Default: true
enabled = true

# The interval in minutes for how often the background cleanup task runs.
# This task purges expired session data from memory to prevent leaks.
# Default: 5
cleanup_interval_minutes = 5

# The maximum number of components that can be updated in a single async response.
# This is a safeguard to prevent a single request from consuming excessive
# server resources or creating an overly large JSON response.
# If you call AxonLive.RegisterComponent more than this number of times,
# the server will raise an ErrG3ALComponentLimitExceeded error.
# Default: 200
max_components_per_response = 200

```

## Session Timeout

The idle timeout for G3AxonLive sessions is not configured directly. Instead, it is derived from the global script timeout to ensure consistency. The formula is:

**Session Idle Timeout = `global.default_script_timeout` x 20**

There is a minimum floor of **30 minutes**.

For example, if `global.default_script_timeout` is set to `90` (seconds), the G3AxonLive session timeout will be `90 * 20 = 1800` seconds, which is 30 minutes. If you increase the script timeout, the session timeout will increase proportionally.

If a session is idle for longer than this duration (i.e., no full page loads and no G3AxonLive async requests are received), the background cleanup task will remove all its associated data from the in-memory store.
