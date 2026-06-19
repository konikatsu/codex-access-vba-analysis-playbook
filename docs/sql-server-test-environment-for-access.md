# Access linked tables and SQL Server test environments

This note summarizes how to choose a local SQL Server test environment when an
Access application uses linked SQL Server tables.

## Summary

If the Access links can be recreated, Docker is often the easiest way to manage
many test databases. Use `server,port` style connection strings such as:

```text
localhost,14333
```

or, if a local host alias is useful:

```text
app-sql,14333
```

If the application must keep a Windows SQL Server named-instance connection such
as:

```text
server-name\instance-name
```

install SQL Server Developer or Express locally as a named instance instead of
using Docker.

## Why named instances are awkward in Docker

SQL Server containers are normally used as one default SQL Server instance per
container. The host maps a TCP port to the container's SQL Server port:

```text
host port 14333 -> container port 1433
```

Clients connect with `server,port`, not `server\instance`.

The `server\instance` syntax depends on SQL Server Browser-style instance name
resolution. Recreating that behavior around Linux SQL Server containers is
possible only with extra plumbing and is usually not worth it for development
and test databases.

## Recommended Docker pattern

Use one container and one named volume per test database or project:

```powershell
docker run `
  -e "ACCEPT_EULA=Y" `
  -e "MSSQL_SA_PASSWORD=<strong-password>" `
  -e "MSSQL_PID=Developer" `
  -p 14333:1433 `
  --name sqlserver-test01 `
  --hostname sqlserver-test01 `
  -v sqlserver-test01-data:/var/opt/mssql `
  -d mcr.microsoft.com/mssql/server:2022-latest
```

Verify the connection:

```powershell
& 'C:\Program Files\SqlCmd\sqlcmd.exe' `
  -S localhost,14333 `
  -U sa `
  -P '<strong-password>' `
  -C `
  -Q "SELECT @@VERSION"
```

The volume is important. If the container is removed without persistent storage,
the database files inside the container are lost.

## Access relink strategy

When using Docker, relink the Access tables to the container endpoint:

```text
Before: server-name\instance-name
After:  localhost,14333
```

or:

```text
After:  app-sql,14333
```

If using a host alias, add the alias to the local hosts file or DNS so it points
to the Docker host. The connection still uses a comma and port, not a backslash
and instance name.

## Decision guide

- Use Docker when the Access linked tables can be relinked to `server,port`.
- Use local SQL Server Developer when many databases need full SQL Server
  features without Express limits.
- Use local SQL Server Express when the environment should resemble a small,
  free, installed SQL Server instance.
- Use LocalDB only for single-user developer tests where services and remote
  style connections are not needed.
- Avoid trying to force Docker SQL Server to behave like a Windows named
  instance unless there is a very specific reason.

## Practical rule

For Access systems, the simplest test strategy is usually:

1. Run SQL Server in Docker on a fixed host port.
2. Restore or create the test database in that container.
3. Recreate the Access linked tables to `localhost,<port>` or a host alias with
   the same port.
4. Keep each project isolated with its own container name, port, and volume.

## References

- Microsoft Learn: [Quickstart: Run SQL Server Linux container images with Docker](https://learn.microsoft.com/en-us/sql/linux/install-upgrade/quickstart-install-docker)
