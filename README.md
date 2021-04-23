# Export-SqlQueryToCsv
This Powershell script will execute a SQL command againsts a SQL database and stream the outputs to a CSV file as the results are returned. Since the resultset is streamed to file, this requires very little RAM and is very performant. For maximum performance, run the script on the database server itself to avoid network latency.

This script should work with Powershell 5+.

## Parameters

| Parameter | Description |
| --------- | ----------- |
| query | The SQL query you want to execute on the server (allows piped strings). |
| output | The absolute or relative path to the CSV file you want to create. If the file exists, it will be overwritten. |
| database | The database to use for the query (you must supply either the database or the dbconnection argument). |
| dbconnection | The full connection string to use for the query (this overrides the datasource and database arguments). |
| datasource | The datasource to use for the query (default: localhost,1433). |
| delim | The delimiter to use for the exported CSV file (default: ,). |
| quote | The character to use to wrap quoted text (default: "). |
| escape | The character to use to escape quoted text characters (default: "). |
| connectionTimeout | The connection timeout to use when executing the query (default: 0; which runs until the query completes). |
| encoding | The encoding to use for the exported file (default: utf-8).|
| dateFormat | The .NET formatting mask to use when rendering DateTime objects (default: "yyyy-MM-dd HH:mm:ss"). |

## Examples

### Simple
The following example takes the query from the first unnamed parameter supplied as an argument and executes it against the server.

```
.\query_to_csv.ps1 -database "my-database-name" -output ".\exported-results.csv" "select * from tableName"
```

### Query Named Parameter
Optionally, you can specify the SQL statement to execute via the `query` parameter:

```
.\query_to_csv.ps1 -query "select * from tableName" -database "my-database-name" -output ".\exported-results.csv"
```

### Query as piped string
Or if your SQL is coming from another script/cmdlet, you can pipe in the string:

```
"select * from tableName" | .\query_to_csv.ps1 -database "my-database-name" -output ".\exported-results.csv"
```
