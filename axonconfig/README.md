# AxonConfig (`axonconfig`)

The `axonconfig` package handles configuration management, environment variables, and project metadata for the AxonASP ecosystem.

## Features

- **Viper Configuration Management**: Centralized loader and schema mapping for `config/axonasp.toml` and `.env` files via `axonconfig.GetConfig()`.
- **Default Fallbacks**: Provides resilient defaults for server, networking, session storage, and runtime VM settings.
- **Project Metadata**: Exposes the `AboutG3pixAxonASP()` function for retrieving versioning, branding, and build information.
