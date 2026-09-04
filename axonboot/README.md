# AxonBoot (`axonboot`)

`axonboot` provides platform-specific early boot initialization logic for Windows environments, particularly when running under Microsoft IIS hosting scenarios (such as FastCGI or HttpPlatformHandler).

## Key Responsibilities

- **Console & Process Setup**: Initializes standard I/O handles and Windows-specific process contexts before any subsystem attempts console output.
- **IIS Integration**: Ensures seamless process lifecycle handling when executed as a worker process spawned by IIS.
- **Early Execution**: Designed to be invoked at the very beginning of `main()` entrypoints before other application logic executes.
