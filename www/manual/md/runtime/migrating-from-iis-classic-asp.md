# Migrating from IIS & Classic ASP to AxonASP: A Deep Architectural Guide

## Overview

Migrating from a legacy Windows Server / IIS environment to AxonASP represents a fundamental paradigm shift in how your Classic ASP applications are hosted, executed, and scaled.

Historically, Classic ASP was deeply coupled with the Windows OS. It ran as an ISAPI extension (`asp.dll`) loaded directly into the IIS worker process (`w3wp.exe`). AxonASP completely severs this dependency. It operates on a modern, application-centric architecture—structurally identical to how modern **.NET Core**, **Node.js**, or **Docker** applications are deployed.

This guide provides a deep technical breakdown of the AxonASP architecture, specifically addressing concerns regarding multi-site hosting, port management, Application Pool isolation, SSL/TLS handling, and the shift away from legacy Windows COM objects.

---

## The App-Centric Execution Model

Because AxonASP is a cross-platform, standalone executable written in Go, it cannot be loaded as a legacy DLL. Furthermore, the native IIS implementation of FastCGI relies on Windows named pipes, which are incompatible with standard FastCGI TCP protocols. Therefore, AxonASP scripts **do not run as an integrated module within IIS**.

Instead, you must provision an isolated, local server instance of AxonASP for each individual application.

### The Golden Rule of Deployment

**AxonASP must never be exposed directly to the public web.** It is a pure ASP execution engine. It intentionally lacks TLS/SSL termination, deep packet inspection, and DDoS mitigation.

Your front-end web server (IIS, Nginx, or Caddy) must remain the public-facing gateway. Its job is to manage the public network layers and securely reverse-proxy dynamic ASP requests to the local AxonASP instances running safely behind the firewall.

---

## Architectural Deep Dive: Hosting Dozens of Sites on One Server

For server administrators accustomed to IIS managing everything seamlessly, a common point of confusion is how to host 50 different websites on standard web ports (80 and 443) and manage FTP access (port 21) if AxonASP requires its own standalone server process.

### 1. Port Routing and the Reverse Proxy

In a multi-site AxonASP architecture, IIS continues to bind exclusively to your public IP addresses on ports 80 and 443. AxonASP **never** touches these ports.

Instead, each website is assigned its own private internal port (e.g., `8801`, `8802`, `8803`) bound exclusively to the `localhost` loopback interface.

When a request arrives at `[https://www.client-site.com/index.asp](https://www.client-site.com/index.asp)`:

1. IIS receives the request on port 443.
2. IIS matches the host header (`[www.client-site.com](https://www.client-site.com)`).
3. Using the **HttpPlatformHandler**, IIS acts as a reverse proxy, forwarding the raw HTTP request internally to `http://localhost:8801` (the AxonASP process dedicated to that specific site).
4. AxonASP compiles and executes the VBScript/JScript, generates the HTML, and hands the response back to IIS.
5. IIS returns the payload to the user over the encrypted TLS connection.

### 2. Static Assets, FTP, and SSL Certificates

Because IIS acts as the master gateway, your surrounding infrastructure remains largely unchanged:

* **FTP Access (Port 21):** FTP remains entirely governed by IIS or your dedicated FTP server. AxonASP is completely unaware of FTP. Users continue to authenticate and upload files to their designated physical paths (e.g., `C:\inetpub\wwwroot\client-site\`).
* **TLS/SSL Certificates:** You do not install certificates into AxonASP. Certificate binding, renewal (e.g., Let's Encrypt / Win-ACME), and TLS handshake negotiation are handled exclusively by IIS.
* **Static File Delivery:** While AxonASP can serve static files, it is architecturally inefficient to do so. Configure IIS to immediately serve `.jpg`, `.css`, `.js`, and `.html` files directly from the disk. Only requests ending in `.asp` (or extensionless routes) should be proxied to AxonASP.

---

## Process Isolation vs. IIS Application Pools

In traditional IIS environments, site isolation is managed via Application Pools. These pools use Windows User Accounts (e.g., `IIS_IUSRS` or `ApplicationPoolIdentity`) to restrict folder access and isolate worker processes (`w3wp.exe`).

AxonASP achieves isolation through **strict process-level boundaries**.

1. **Shared-Nothing Architecture:** Every application runs its own distinct instance of `axonasp-http.exe`. If Site A writes infinite loops and exhausts its memory, the process may crash, but Site B, running in a completely separate OS process, remains 100% unaffected.
2. **State Management:** Classic ASP `Application` variables and the `global.asa` lifecycle exist purely within the RAM of a single process.
3. **The `processesPerApplication` Rule:** Because of this in-memory state, **you must configure IIS to spawn exactly 1 process per application**. If IIS spawns multiple AxonASP processes for a single site, user sessions will fracture, and `Application` state will become desynchronized, leading to unpredictable application behavior.
4. **Security Boundaries:** You can still utilize Windows ACLs. Simply configure the IIS HttpPlatformHandler to launch the `axonasp-http.exe` process under a restricted Windows user account, ensuring the process cannot read the physical directory of a neighboring website.

---

## Reconfiguration: URL Rewriting

Administrators heavily rely on URL rewriting for SEO-friendly URLs. In the AxonASP architecture, this responsibility can be split depending on your needs:

* **Global/Infrastructure Rewriting (IIS side):** Heavy lifting, such as forcing `HTTP -> HTTPS` redirects, appending `www.`, or blocking malicious IP ranges, should remain in the IIS URL Rewrite module. This stops bad traffic before it ever reaches the AxonASP engine.
* **Application Routing (AxonASP side):** AxonASP includes its own native parser for a simplified `web.config` file. If you use a "Front Controller" pattern (routing all traffic to `index.asp?route=...`), you should use the AxonASP `web.config`. This ensures that your application logic remains entirely portable. If you ever migrate from Windows/IIS to Linux/Docker, your routing rules move with your application automatically.

---

## Extending the Engine: The Post-COM Era

A significant hurdle when migrating from Classic ASP is the reliance on legacy Windows COM objects (e.g., ADODB, CDO, MSXML2) registered via `regsvr32`.

Because AxonASP is written in Go and designed for cross-platform execution, **it actively bypasses the Windows OS Registry and legacy COM integration.**

### The Challenge with GoOLE

We recognize that for developers unfamiliar with Go, utilizing packages like `GoOLE` to wrap and execute legacy Windows COM components is highly complex and time-consuming. Re-implementing thousands of lines of COM bridges defeats the purpose of modernization.

### The AxonASP Native Routing Solution

Instead of relying on OS-level COM objects, AxonASP intercepts the `Server.CreateObject("ProgID")` call natively.
When your legacy code requests an object, the AxonASP Virtual Machine instantly maps that ProgID to an ultra-fast, pre-compiled Go struct running directly inside the VM.

* **For Legacy Compatibility:** AxonASP ships with native emulations of critical objects like `Scripting.FileSystemObject`, `Scripting.Dictionary`, and `ADODB` (which connects to real modern databases like PostgreSQL or MSSQL).
* **For Modern Capabilities:** You should transition away from legacy COM objects and utilize the built-in Axon Enterprise Libraries (`G3JSON`, `G3HTTP`, `G3DB`, `G3Crypto`, etc.). These execute with zero-allocation overhead and require no OS registration.
* **Future Extensibility:** We understand the community needs a simpler path to custom extensions. In a future update, we will provide a streamlined, heavily documented boilerplate example demonstrating exactly how to create a custom Go library, attach it to a ProgID, and compile it into AxonASP without needing deep GoLang expertise.

---

## Implementation Example: Running via HttpPlatformHandler

To properly orchestrate AxonASP within IIS, you must use the **HttpPlatformHandler v1.2+** module. This module instructs IIS to act as a process manager. It will automatically launch `axonasp-http.exe`, monitor its health, and dynamically assign a port for internal reverse proxying.

1. Isolate your application in a specific directory (e.g., `C:\Sites\App1\`).
2. Ensure `axonasp-http.exe`, the `axonasp.toml` config file, the `iis-http.cmd` wrapper, and your `www` folder reside there.
3. Add the following `web.config` to the root of the IIS site:

```xml
<?xml version="1.0" encoding="UTF-8"?>
<configuration>
    <system.webServer>
        <handlers>
            <add name="httpPlatformHandler" path="*" verb="*" modules="httpPlatformHandler" resourceType="Unspecified" />
        </handlers>
        
        <!-- 
          IIS acts as the process manager here. 
          %HTTP_PLATFORM_PORT% is a dynamic variable injected by IIS. 
          It ensures this specific AxonASP instance gets a unique, conflict-free internal port.
        -->
        <httpPlatform processPath="C:\Sites\App1\iis-http.cmd"
                      arguments="--server.server_port %HTTP_PLATFORM_PORT%"
                      stdoutLogEnabled="true"
                      stdoutLogFile="C:\Sites\App1\temp\axonasp.log"
                      startupTimeLimit="5"
                      processesPerApplication="1"> <!-- CRITICAL: Must be 1 to preserve global.asa and session state -->
            <environmentVariables>
                <environmentVariable name="SERVER_PORT" value="%HTTP_PLATFORM_PORT%" />
            </environmentVariables>
        </httpPlatform>
    </system.webServer>
</configuration>

```

By relying on `%HTTP_PLATFORM_PORT%`, you do not need to manually track whether Site A is on port `8801` and Site B is on `8802`. IIS assigns an available port dynamically, boots the AxonASP process on that port, and instantly begins proxying the traffic.

## The Containerization Horizon (Docker)

Once you understand that AxonASP is simply a self-contained executable processing HTTP requests, the ultimate migration path becomes clear. You are no longer bound to IIS.

You can package your ASP code, your `axonasp.toml` configuration, and the Linux version of `axonasp-http` into a Docker container. Using an ingress controller (like Nginx or Traefik), you can dynamically route traffic to hundreds of isolated ASP containers across a cluster, achieving enterprise-grade scalability for legacy codebase.