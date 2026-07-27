<%@ Language="JScript" %>
<%
function OpenCmsDb() {
    var db = Server.CreateObject("G3DB");
    var dsn = G3PIX_DB_FILE;
    if (dsn.indexOf("?") < 0) {
        dsn += "?_busy_timeout=6000";
    }
    if (!db.Open("sqlite", dsn)) {
        return null;
    }
    return db;
}

function CloseCmsDb(db) {
    if (db != null) {
        db.Close();
        db = null;
    }
}

function HasDbColumn(db, tableName, columnName) {
    var rs = db.Query("PRAGMA table_info(" + tableName + ")");
    var exists = false;

    if (rs != null && rs != "") {
        while (!rs.EOF) {
            if (String(rs("name")) == columnName) {
                exists = true;
                break;
            }
            rs.MoveNext();
        }
        rs.Close();
    }

    return exists;
}

function EnsureDbColumn(db, tableName, columnName, columnDefinition) {
    if (!HasDbColumn(db, tableName, columnName)) {
        db.Exec("ALTER TABLE " + tableName + " ADD COLUMN " + columnDefinition);
    }
}

function EnsureSchemaAndSeed() {
    EnsureDataFolder();
    EnsureUploadFolder();

    var db = OpenCmsDb();
    if (db == null) {
        return false;
    }

    db.Exec("CREATE TABLE IF NOT EXISTS users (id INTEGER PRIMARY KEY AUTOINCREMENT, username TEXT NOT NULL UNIQUE, password_hash TEXT NOT NULL, display_name TEXT NOT NULL, is_active INTEGER NOT NULL DEFAULT 1, created_at TEXT NOT NULL DEFAULT (datetime('now'))) ");
    db.Exec("CREATE TABLE IF NOT EXISTS pages (id INTEGER PRIMARY KEY AUTOINCREMENT, slug TEXT NOT NULL UNIQUE, title TEXT NOT NULL, excerpt TEXT NOT NULL DEFAULT '', content TEXT NOT NULL DEFAULT '', status TEXT NOT NULL DEFAULT 'draft', locale TEXT NOT NULL DEFAULT 'en', is_home INTEGER NOT NULL DEFAULT 0, page_type TEXT NOT NULL DEFAULT 'page', parent_id INTEGER NOT NULL DEFAULT 0, sort_order INTEGER NOT NULL DEFAULT 0, hero_badge TEXT NOT NULL DEFAULT '', hero_title TEXT NOT NULL DEFAULT '', hero_subtitle TEXT NOT NULL DEFAULT '', hero_content TEXT NOT NULL DEFAULT '', hero_primary_label TEXT NOT NULL DEFAULT '', hero_primary_url TEXT NOT NULL DEFAULT '', hero_secondary_label TEXT NOT NULL DEFAULT '', hero_secondary_url TEXT NOT NULL DEFAULT '', home_mode TEXT NOT NULL DEFAULT 'all', home_section_title TEXT NOT NULL DEFAULT '', seo_title TEXT NOT NULL DEFAULT '', meta_description TEXT NOT NULL DEFAULT '', created_at TEXT NOT NULL DEFAULT (datetime('now')), updated_at TEXT NOT NULL DEFAULT (datetime('now'))) ");
    db.Exec("CREATE TABLE IF NOT EXISTS menus (id INTEGER PRIMARY KEY AUTOINCREMENT, title TEXT NOT NULL, page_slug TEXT NOT NULL DEFAULT '', url TEXT NOT NULL DEFAULT '', target TEXT NOT NULL DEFAULT '_self', sort_order INTEGER NOT NULL DEFAULT 0, is_visible INTEGER NOT NULL DEFAULT 1, locale TEXT NOT NULL DEFAULT 'en', created_at TEXT NOT NULL DEFAULT (datetime('now')), updated_at TEXT NOT NULL DEFAULT (datetime('now'))) ");
    db.Exec("CREATE TABLE IF NOT EXISTS media (id INTEGER PRIMARY KEY AUTOINCREMENT, file_name TEXT NOT NULL, relative_path TEXT NOT NULL, mime_type TEXT NOT NULL DEFAULT '', size_bytes INTEGER NOT NULL DEFAULT 0, uploaded_by TEXT NOT NULL DEFAULT '', created_at TEXT NOT NULL DEFAULT (datetime('now'))) ");
    db.Exec("CREATE TABLE IF NOT EXISTS settings (setting_key TEXT PRIMARY KEY, setting_value TEXT NOT NULL) ");
    db.Exec("CREATE TABLE IF NOT EXISTS js_snippets (id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT NOT NULL, code TEXT NOT NULL, is_active INTEGER NOT NULL DEFAULT 1, sort_order INTEGER NOT NULL DEFAULT 0, created_at TEXT NOT NULL DEFAULT (datetime('now')), updated_at TEXT NOT NULL DEFAULT (datetime('now'))) ");

    EnsureDbColumn(db, "users", "role", "role TEXT NOT NULL DEFAULT 'admin'");
    EnsureDbColumn(db, "users", "must_change_password", "must_change_password INTEGER NOT NULL DEFAULT 1");
    EnsureDbColumn(db, "pages", "seo_title", "seo_title TEXT NOT NULL DEFAULT ''");
    EnsureDbColumn(db, "pages", "meta_description", "meta_description TEXT NOT NULL DEFAULT ''");
    EnsureDbColumn(db, "pages", "page_type", "page_type TEXT NOT NULL DEFAULT 'page'");
    EnsureDbColumn(db, "pages", "parent_id", "parent_id INTEGER NOT NULL DEFAULT 0");
    EnsureDbColumn(db, "pages", "sort_order", "sort_order INTEGER NOT NULL DEFAULT 0");
    EnsureDbColumn(db, "pages", "hero_badge", "hero_badge TEXT NOT NULL DEFAULT ''");
    EnsureDbColumn(db, "pages", "hero_title", "hero_title TEXT NOT NULL DEFAULT ''");
    EnsureDbColumn(db, "pages", "hero_subtitle", "hero_subtitle TEXT NOT NULL DEFAULT ''");
    EnsureDbColumn(db, "pages", "hero_content", "hero_content TEXT NOT NULL DEFAULT ''");
    EnsureDbColumn(db, "pages", "hero_primary_label", "hero_primary_label TEXT NOT NULL DEFAULT ''");
    EnsureDbColumn(db, "pages", "hero_primary_url", "hero_primary_url TEXT NOT NULL DEFAULT ''");
    EnsureDbColumn(db, "pages", "hero_secondary_label", "hero_secondary_label TEXT NOT NULL DEFAULT ''");
    EnsureDbColumn(db, "pages", "hero_secondary_url", "hero_secondary_url TEXT NOT NULL DEFAULT ''");
    EnsureDbColumn(db, "pages", "home_mode", "home_mode TEXT NOT NULL DEFAULT 'all'");
    EnsureDbColumn(db, "pages", "home_section_title", "home_section_title TEXT NOT NULL DEFAULT ''");

    db.Exec("UPDATE pages SET page_type = 'home' WHERE is_home = 1 AND page_type <> 'home'");

    var rsUser = db.Query("SELECT id FROM users WHERE username = ?", "admin");
    var hasDefaultUser = false;
    if (rsUser != null && rsUser != "") {
        if (!rsUser.EOF) {
            hasDefaultUser = true;
        }
        rsUser.Close();
    }

    if (!hasDefaultUser) {
        var crypto = Server.CreateObject("G3CRYPTO");
        var passwordHash = crypto.HashPassword("123change");
        db.Exec("INSERT INTO users (username, password_hash, display_name, is_active, role, must_change_password) VALUES (?, ?, ?, ?, ?, ?)", "admin", passwordHash, "Admin", 1, "admin", 1);
        crypto = null;
    }

    var rsHome = db.Query("SELECT id FROM pages WHERE is_home = ?", 1);
    var hasHome = false;
    if (rsHome != null && rsHome != "") {
        if (!rsHome.EOF) {
            hasHome = true;
        }
        rsHome.Close();
    }

    if (!hasHome) {
        db.Exec("INSERT INTO pages (slug, title, excerpt, content, status, locale, is_home, page_type, home_mode, home_section_title, seo_title, meta_description) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
            "home",
            "Welcome to G3Pix CMS",
            "Edit this homepage in the Pages section.",
            "<p>Build your homepage with reusable blocks like [g3pix-part slug=\"header\"] and [g3pix-part slug=\"footer\"].</p>",
            "published",
            "en",
            1,
            "home",
            "all",
            "Latest content",
            "",
            ""
        );
    }

    var rsMenu = db.Query("SELECT id FROM menus");
    var hasMenuItems = false;
    if (rsMenu != null && rsMenu != "") {
        if (!rsMenu.EOF) {
            hasMenuItems = true;
        }
        rsMenu.Close();
    }

    if (!hasMenuItems) {
        db.Exec("INSERT INTO menus (title, page_slug, url, target, sort_order, is_visible, locale) VALUES (?, ?, ?, ?, ?, ?, ?)", "Home", "home", "", "_self", 1, 1, "en");
    }

    db.Exec("INSERT OR IGNORE INTO settings (setting_key, setting_value) VALUES (?, ?)", "site_title", "G3Pix CMS");
    db.Exec("INSERT OR IGNORE INTO settings (setting_key, setting_value) VALUES (?, ?)", "default_locale", "en");
    db.Exec("INSERT OR IGNORE INTO settings (setting_key, setting_value) VALUES (?, ?)", "custom_css", "");
    db.Exec("INSERT OR IGNORE INTO settings (setting_key, setting_value) VALUES (?, ?)", "site_logo_url", "/logo_square.svg");
    db.Exec("INSERT OR IGNORE INTO settings (setting_key, setting_value) VALUES (?, ?)", "footer_text", "G3Pix CMS on AxonASP");
    db.Exec("INSERT OR IGNORE INTO settings (setting_key, setting_value) VALUES (?, ?)", "show_search", "1");
    db.Exec("INSERT OR IGNORE INTO settings (setting_key, setting_value) VALUES (?, ?)", "show_home_menu", "1");
    db.Exec("INSERT OR IGNORE INTO settings (setting_key, setting_value) VALUES (?, ?)", "show_login_menu", "1");
    db.Exec("INSERT OR IGNORE INTO settings (setting_key, setting_value) VALUES (?, ?)", "show_admin_menu", "1");
    db.Exec("INSERT OR IGNORE INTO settings (setting_key, setting_value) VALUES (?, ?)", "show_sitemap_menu", "1");

    CloseCmsDb(db);
    return true;
}
%>