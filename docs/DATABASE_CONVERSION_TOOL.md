# AxonASP Database Conversion Tool

## Overview

The **AxonASP Database Conversion Tool** is a built-in module designed to facilitate the migration of legacy Microsoft Access databases (`.mdb` and `.accdb` formats) to modern database systems supported by AxonASP, including SQLite, MySQL, PostgreSQL, and MS SQL Server.

This tool provides a user-friendly, step-by-step wizard interface to handle file uploads, table selection, schema mapping, and data migration.

## Accessing the Tool

The tool is available at the following URL on your local AxonASP server:
`http://localhost:4050/database-convert/`

## Key Features

- **Multi-Target Support**: Migrate data to SQLite, MySQL, PostgreSQL, or Microsoft SQL Server.
- **Smart Schema Mapping**: Automatically maps Access data types to the most appropriate types in the target database (INTEGER, TEXT, FLOAT, DATETIME).
- **Flexible Connection Modes**:
    - **System Default**: Uses database credentials and settings defined in your `.env` file.
    - **Manual Configuration**: Allows entering specific server address, port, username, password, and database name.
- **Asynchronous Processing**: Uses AJAX-based worker scripts to perform migrations table-by-table, providing real-time progress feedback.
- **Transactional Integrity**: Uses database transactions during the record copying phase to ensure data consistency and improved performance.
- **Table Selection**: Choose specifically which tables to migrate from the source database.

## Technical Requirements

### Source (Reading Access)
- **Windows**: Requires Microsoft Access Database Engine (JET or ACE OLEDB provider) installed on the host machine.
- **Linux/macOS**: Direct reading of `.mdb`/`.accdb` files is currently only supported via the Windows OLEDB bridge. For cross-platform scenarios, it is recommended to run the conversion on a Windows host or use pre-converted SQLite files.

### Target (Writing)
- Requires the appropriate database driver to be enabled in AxonASP (SQLite, MySQL, PostgreSQL, and MS SQL Server are supported natively).

## Step-by-Step Guide

### Step 1: Source & Target Configuration
1. Upload your `.mdb` or `.accdb` file.
2. Select the **Target Type** (e.g., MySQL).
3. Choose the **Connection Mode**:
    - **Use System Default**: Pre-fills settings from your `.env` file.
    - **Manual Configuration**: Manually enter host, port, credentials, and database name.

### Step 2: Table Selection
The tool will analyze the Access database and list all available tables. Select the checkboxes for the tables you wish to migrate.

### Step 3: Migration Process
Click **Next** to begin the migration. The tool will:
1. Create the tables in the target database (dropping existing ones with the same name).
2. Copy records in batches using transactions.
3. Show progress for each table in the log window.

### Step 4: Completion
Once finished, a success message will display the target database details. You can then close the wizard and begin using your data in the new format.

## Implementation Details

- **Controller**: `www/database-convert/wizard.asp`
- **Worker**: `www/database-convert/import_worker.asp`
- **Libraries Used**:
    - `G3FileUploader`: For secure handling of the database file upload.
    - `ADOX.Catalog`: To inspect the schema of the Access database.
    - `ADODB.Connection`: For managing connections to both source and target databases.
    - `ADODB.Recordset`: For iterating through source records.

## Security Considerations

- Access to the `database-convert` folder should be restricted in production environments.
- Uploaded files are temporarily stored in `temp/uploads/` and should be cleaned up according to your server's maintenance policy.
- The tool uses parameterized queries where applicable during the migration phase to ensure data integrity.

## See Also

- [ADODB_IMPLEMENTATION.md](ADODB_IMPLEMENTATION.md)
- [ACCESS_DATABASE_SUPPORT.md](ACCESS_DATABASE_SUPPORT.md)
- [G3DB_IMPLEMENTATION.md](G3DB_IMPLEMENTATION.md)
