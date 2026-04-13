# G3DB Properties

## Overview
This page provides a summary of the properties available in the **G3DB** library for inspecting the state of a database connection.

## Property List

- **Driver**: Read-only. Returns the canonical name of the database driver being used (e.g., `mysql`, `postgres`, `sqlite`).
- **DSN**: Read-only. Returns the Data Source Name (connection string) used to establish the connection.
- **IsOpen**: Read-only. Returns a **Boolean** indicating whether the database connection pool is currently active and open.
- **LastError**: Read-only. Returns the last error message recorded by the database connection object.

## Remarks
- Properties are read-only and will raise a runtime error if an assignment is attempted.
- Property access is efficient and does not trigger new database operations.
