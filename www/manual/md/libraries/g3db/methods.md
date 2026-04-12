# Methods

## Overview

This page lists methods exposed by G3DB.

## Method List
- Begin: Begins a transactional or grouped operation scope.
- BeginTx: Begins Tx for grouped operations.
- Close: Closes the current resource and releases handles.
- Exec: Executes a command and returns status or output.
- GetError: Gets Error from the G3DB library.
- Open: Opens a resource for subsequent operations.
- OpenFromEnv: Opens From Env for subsequent operations.
- Prepare: Prepares a statement for efficient repeated execution.
- Query: Executes a query and returns results.
- QueryRow: Executes a query for Row and returns results.
- SetConnMaxIdleTime: Sets Conn Max Idle Time for the G3DB library.
- SetConnMaxLifetime: Sets Conn Max Lifetime for the G3DB library.
- SetMaxIdleConns: Sets Max Idle Conns for the G3DB library.
- SetMaxOpenConns: Sets Max Open Conns for the G3DB library.
- Stats: Returns structured runtime information for the current context.

## Remarks

- Method names are case-insensitive.
- Validate input types and return values in production code.
- Use Set when assigning object return values.
