# AxonASP FPM (`fpm`)

The `fpm` package implements the FastCGI Process Manager (FPM) for AxonASP, designed to manage, monitor, and scale FastCGI worker processes.

## Key Responsibilities

- **Worker Pool Management**: Dynamically spawns, monitors, and recycles FastCGI child worker instances based on workload and configuration.
- **Process Supervision**: Handles health checks, graceful restarts, and automatic recovery of crashed worker processes.
- **IPC & Socket Orchestration**: Coordinates listening sockets (TCP or Unix/named pipes) across multiple worker instances for load balancing.
