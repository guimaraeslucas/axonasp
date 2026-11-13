When writing or updating documentation for AxonASP, follow these guidelines:

* **Authoring Guidelines:** Before generating or updating any manual page, you MUST read and strictly follow the guidelines in `./www/manual/md/authoring/write-manual-pages.md`.
* **When to Write:** Create or update manual pages when new libraries, methods, properties, or significant features are added. Always ensure that documentation is up-to-date with the latest implementation.
* **Location:** `www\manual\md\` for markdown content, `www\manual\menu.md` for the navigable menu.
* **Format:** Follow Microsoft Writing Style Guide (action-oriented titles, brief overview, prerequisites, code examples, extra information for how the code works and API references). *Don't create markdown links inside the content pages. Use markdown links only in `menu.md` for navigation.*
* **Style:** Active voice, simple language, scannable lists, and bold text. NO EMOJIS.
* **Branding:** Use AxonASP branding. DO NOT use Microsoft names/logos for our functions.
* **Menu:** Always update `www\manual\menu.md` (using a nested bulleted list) after creating new docs.
* **Code Examples:** Always provide a complete, copy-pastable code example that can be run immediately. Include the necessary script tags and initialization for AxonLive if applicable.
* **Documentation of New Errors:** When a new error code is implemented in the libraries or VM, add it to the `www\manual\md\runtime\axonasp-error-codes.md` page with its number, description, and possible causes.
* **Documentation of New Configurations:** When a new configuration option is added to `config/axonasp.toml`, add it to the `www\manual\md\runtime\configuration.md` page with its name, description, and default value.
* **Documentation of New Libraries/Functions:** When a new library or function is added, create a new markdown page in `www\manual\md\libraries\` with its name and document its usage, parameters, return values, and examples.
* **Nomenclature:** Don't use JScript in documentation, use JavaScript instead, as it is more widely recognized by users. Only use "JScript" when specifically referring to the engine or compatibility mode.