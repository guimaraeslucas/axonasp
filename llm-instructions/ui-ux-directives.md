# UI/UX Directives for AxonASP/G3Pix Interfaces

* **Design Direction:** Modern, clean, and product-like interface inspired by enterprise dashboards. Prioritize readability, spacing rhythm, and visual hierarchy. We must keep a ressemble of 
* **Constraints:** NO FRAMEWORKS (No Bootstrap/Tailwind). Use vanilla HTML5, JS, and the shared stylesheet `./www/css/axonasp.css`. If a new reusable rule is required, add it to this file. Do not embed large inline style blocks.
* **Theme System:** Support both light and dark modes using the global token model already defined in `:root`, `@media (prefers-color-scheme: dark)`, and `:root[data-theme="light|dark"]` in `axonasp.css`.
* **Token Usage:** Always consume CSS custom properties from `axonasp.css`. Never hardcode color palettes in markup. Prefer semantic tokens such as:
    * `--bg`, `--bg-elevated`, `--bg-soft`
    * `--text`, `--text-muted`
    * `--border`, `--surface-hover`
    * `--accent`, `--accent-strong`, `--accent-soft`
    * `--success`, `--danger`, `--warning`
* **Typography:** Use the current global typography stack from `axonasp.css` (IBM Plex Sans / IBM Plex Mono fallbacks). Keep line length and spacing readable for technical content.
* **Geometry & Surfaces:** Respect the active geometry tokens (`--radius-*`) and elevation tokens (`--shadow`, `--shadow-soft`). Use cards and panels to group related information.
* **Core Components:** Use existing component classes from `axonasp.css` whenever possible (`.btn*`, `.card*`, `.alert*`, `.badge*`, `.pill*`, `.table-wrap`, `.grid-*`, `.window*`, `.treeview`, `.actions-row`, `.info-banner`, `.cta-panel`). Extend by adding reusable classes in CSS, not per-page inline overrides.
* **Layout Rules:** Keep shell consistency with `#header`, `#main-container`, `#content`, and `#status-bar` conventions already used across the app. Use responsive spacing and avoid fixed-width layouts that break on smaller screens.
* **Page-Level Styling:** For page-specific behavior, scope styles with clear page classes (e.g., `.manual-page`, `.project-builder-page`, `.jsapi-page`, `.wasm-playground-page`) in `axonasp.css`.
* **Branding:** Use AxonASP/G3Pix branding assets and naming only. Do not reference Microsoft/MSDN branding.
* **Quality Guardrails:** Ensure visual consistency across all pages (manual, project-builder, default, API demos, error pages, WASM playground). Prefer class-based state changes (e.g., loading/success/error) over direct inline style mutations in JS.
- **Keep queries focused and generic**: Describe the user-visible capability (`<verb> + <UI/platform object>`, optionally paired with the target HTML/CSS/JS feature name if known). Omit framework/library names (`React`, `Vue`, `Tailwind`, etc.) and app/domain nouns (product, route, or component/file names). Split multi-feature tasks into separate `search` calls.
- **Name implied UI primitives**: When the user's prompt expresses intent without naming the UI primitive (e.g. "confirm first", "delete with confirmation", "pick a date"), translate the domain intent into the implied UI component or platform mechanism (e.g. search `"confirmation modal dialog"` or `"confirm dialog"` rather than `"confirm destructive action"`, or `"date picker input"` rather than `"pick a date"`).