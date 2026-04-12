# Database Convert Tool

## Overview

The Database Convert tool is a built-in AxonASP web application located at `./www/database-convert/` that provides a wizard-style interface for migrating Microsoft Access databases to other database formats supported by AxonASP. The tool reads the source Access database file, inspects its schema and data using ADODB, and writes the content to the selected target.

Access the tool at: `http://localhost:8801/database-convert/`

## Supported Conversions

| Source | Target | Status |
|--------|--------|--------|
| Microsoft Access (.mdb / .accdb) | SQLite | Supported |
| Microsoft Access (.mdb / .accdb) | MySQL | Supported |
| Microsoft Access (.mdb / .accdb) | PostgreSQL | Supported |
| Microsoft Access (.mdb / .accdb) | Microsoft SQL Server | Supported |
| Microsoft Access (.mdb / .accdb) | Oracle | Not supported |

> **Note:** Conversion to Oracle is not currently supported.

## How It Works

The conversion runs as an ASP wizard with multiple steps:

1. **Upload step:** The user uploads a `.mdb` or `.accdb` Access database file.
2. **Target selection step:** The user selects the destination database type and provides connection details.
3. **Conversion step:** The worker script (`import_worker.asp`) opens the source database using ADODB with the ACE OLE DB or Jet OLE DB provider, reads each table's schema, creates equivalent tables in the target database, and inserts all rows in batches.

The tool uses:
- `ADODB.Connection` to open both the source and the target database.
- `G3FILES` for file handling and temporary storage of the uploaded database.
- A target-specific connection string built from the connection parameters provided by the user.

## Prerequisites

- **Windows only:** The source database connection requires the Microsoft Access Database Engine (ACE OLE DB 12.0) or the older Jet OLE DB 4.0 provider. These are Windows-only COM providers.
- The target database must be reachable from the machine running AxonASP.
- For SQLite: the `SQLITE_PATH` environment variable must point to the destination `.db` file.
- For MySQL, PostgreSQL, or SQL Server: valid host, port, database name, username, and password are required.

## Identifier Quoting

The tool applies correct identifier quoting for each target dialect:

| Target | Quoting style |
|--------|---------------|
| SQLite | `[brackets]` |
| SQL Server | `[brackets]` |
| MySQL | `` `backticks` `` |
| PostgreSQL | `"double quotes"` |

## Type Mapping

ADO field types from Access are mapped to type-appropriate SQL column definitions in the target schema. Unsupported or unknown ADO types fall back to a generic text type.

## Running the Tool

1. Start the AxonASP HTTP server.
2. Open a browser and navigate to `http://localhost:8801/database-convert/`.
3. Follow the wizard steps: upload the Access database, select the target, enter connection details, and start the conversion.

## Remarks

- The tool is a development and administration utility. It is not intended for production-facing traffic. Restrict access to this path in production environments using web.config rewrite rules or network-level controls.
- Large databases with many rows may take time to convert. The conversion runs synchronously within the ASP request.
- The Access database file is temporarily stored on the server during conversion. It is removed after the conversion completes.
- Primary keys, indexes, and foreign key constraints from Access are not automatically recreated in the target schema. Only table structure and data are migrated.
