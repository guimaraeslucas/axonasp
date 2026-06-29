# IIS Support

## Overview

Due to how IIS handles FastCGI (named pipes) on Windows, which is incompatible with the GoLang implementation, *it is not possible* to directly integrate the `axonasp-fastcgi.exe` binary into IIS.

Therefore, as recommended by Microsoft for running other engines like Python, AxonASP must be integrated via `httpPlatformHandler`. Another option is using the IIS reverse proxy. In this case, you must start AxonASP separately and ensure continuous execution (you can use a service implementation). The target address must be set to `http://localhost:8801` (modify the port if necessary). The `httpPlatformHandler` implementation is recommended, as it automatically manages the server lifecycle. For further instructions on configuring a reverse proxy, check the documentation for reverse proxy setup or visit https://learn.microsoft.com/en-us/iis/extensions/url-rewrite-module/reverse-proxy-with-url-rewrite-v2-and-application-request-routing.

We provide the `iis-http.cmd` file in both the Windows installation and the project root directory. You must define it in the `processPath="C:\axonasp\iis-http.cmd"` attribute. **Warning:** If AxonASP is not installed in the default directory (`C:\axonasp\`), you must modify the file to point to the correct directory. Otherwise, the system will fail to locate the configuration file, and script execution will fail silently.

Because IIS does not allocate a real console during server execution, console messages via `stdoutLogEnabled="true"` will always yield a blank `stdoutLogFile`. This is a strict limitation caused by the integrated C-supported SQLite library, which crashes the server if started without an attached console. AxonASP error and console messages will continue to be written to the location defined in the `axonasp.toml` configuration file.

## Installation

Before running AxonASP on IIS, you must install the IIS HttpPlatformHandler v1.2 module or higher, available at: https://www.iis.net/downloads/microsoft/httpplatformhandler

## Configuration

Following the installation of the HttpPlatformHandler module, you must configure your IIS `web.config` as follows:

```xml
<?xml version="1.0" encoding="UTF-8"?>
<configuration>
    <system.webServer>
        <handlers>
            <add name="httpPlatformHandler" path="*" verb="*" modules="httpPlatformHandler" resourceType="Unspecified" />
        </handlers>
        <httpPlatform processPath="C:\axonasp\iis-http.cmd"
                      arguments="--server.server_port %HTTP_PLATFORM_PORT%"
                      stdoutLogEnabled="true"
                      stdoutLogFile="C:\axonasp\temp\axonasp.log"
                      startupTimeLimit="5"
                      processesPerApplication="1">
            <environmentVariables>
                <environmentVariable name="SERVER_PORT" value="%HTTP_PLATFORM_PORT%" />
            </environmentVariables>
        </httpPlatform>
    </system.webServer>
</configuration>
```
**Warning:** Ensure the IIS user, typically defined as `IIS_IUSRS`, has read/write permissions to the AxonASP folders and files. This is required for writing to the `temp` directory and for executing both `iis-http.cmd` and `axonasp-http.exe`. Without proper permissions, the process will fail silently and return a 500 Internal Server Error.

You can append the `--config.config_file` flag to the `arguments` attribute in the `web.config` to point to a specific, separate configuration file for each IIS site (e.g. `arguments="--server.server_port %HTTP_PLATFORM_PORT% --config.config_file C:\websites\site_a\axonasp.toml"`). This enables you to serve multiple distinct websites and applications from the same single, shared AxonASP installation directory, entirely eliminating the need to duplicate and copy-paste the server executables into each website's physical directory path.

The executable that IIS will proxy traffic to is `axonasp-http.exe`. Consequently, your ASP files, `global.asa`, AxonASP `web.config`, and overall infrastructure must reside within the directory defined in the configuration file. IIS acts solely as a reverse proxy starting the executable. For proper isolation, we recommend that each AxonASP Application/Site is configured with its own `.toml` configuration file using the custom config path argument. Also, set the `clean_sessions_on_startup` in the configuration file to prevent sessions being lost if IIS terminates the application process, or users may experience unexpected behavior when using features that require persistent sessions. For better performance, you can also set `clean_cache_on_startup` to `false`.


## Limitations
While IIS natively supports Classic ASP, it is highly resource-intensive. If possible, consider migrating to a lighter reverse proxy setup (e.g., Nginx, Caddy).

The Microsoft IIS implementation spawns a FastCGI process with named pipes for each individual request. This contrasts with Nginx/Apache/Caddy architectures, which start a single persistent process and route requests via TCP. This IIS behavior can cause anomalous execution state, particularly regarding session persistence and Application-level memory data.

Because we recommend isolating each application in its own server process due to `global.asa` scoping rules, and because supporting IIS FastCGI directly would require a complex architectural rewrite, native IIS FastCGI support will not be implemented at this time. For this same reason, you must **not** set `processesPerApplication` to a value greater than 1, as it will cause state fragmentation and anomalous behavior.