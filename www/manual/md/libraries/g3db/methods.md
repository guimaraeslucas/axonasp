# G3DB Methods

## Overview
This page provides a summary of the methods available in the **G3DB** library for database operations in AxonASP.

## Method List

- **Begin**: Starts a standard database transaction and returns a **G3DBTransaction** object.
- **BeginTx**: Starts a transaction with optional timeout (seconds) and read-only status. Returns a **G3DBTransaction**.
- **Close**: Shuts down the database connection pool and releases all resources.
- **Exec**: Executes a non-query command (e.g., INSERT, UPDATE, DELETE). Returns a **G3DBResult** object.
- **GetError**: Retrieves the most recent error message encounter by the connection object.
- **Open**: Establishes a connection to the specified database using the requested driver and connection string.
- **OpenFromEnv**: Establishes a connection using parameters from the application configuration.
- **Prepare**: Compiles a prepared SQL statement. Returns a **G3DBStatement** object.
- **Query**: Executes a SQL SELECT statement and returns a forward-only **G3DBResultSet**.
- **QueryRow**: Executes a SQL SELECT statement and returns a **G3DBRow** for single-row results.
- **SetConnMaxIdleTime**: Configures the duration (seconds) a connection can remain idle before being closed.
- **SetConnMaxLifetime**: Configures the maximum duration (seconds) a connection can be reused in the pool.
- **SetMaxIdleConns**: Configures the maximum number of idle connections allowed in the pool.
- **SetMaxOpenConns**: Configures the maximum number of concurrent open connections allowed.
- **Stats**: Returns a **Scripting.Dictionary** containing current database connection pool statistics.

## Remarks
- All methods are case-insensitive.
- Methods that execute SQL support parameterized queries with `?` placeholders.
- Methods returning objects (like **Query**, **Begin**, or **Prepare**) must be assigned using the `Set` keyword.
