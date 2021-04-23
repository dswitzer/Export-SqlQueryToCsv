#requires -version 5

<#
.SYNOPSIS
  Executes a SQL query and generates a CSV file from the resultset.
.DESCRIPTION
  This script will execute a SQL command againsts a SQL database and stream the outputs
	to a CSV file as the results are returned. Since the resultset is streamed to file,
	this requires very little RAM and is very performant. For maximum performance, run the
	script on the database server itself to avoid network latency.
.PARAMETER query
    The SQL query you want to execute on the server (allows piped strings).
.PARAMETER output
    The absolute or relative path to the CSV file you want to create. If the file exists,
		it will be overwritten.
.PARAMETER database
    The database to use for the query (you must supply either the database or the dbconnection argument).
.PARAMETER dbconnection
    The full connection string to use for the query (this overrides the datasource and database arguments).
.PARAMETER datasource
    The datasource to use for the query (default: localhost,1433).
.PARAMETER delim
    The delimiter to use for the exported CSV file (default: ,).
.PARAMETER quote
    The character to use to wrap quoted text (default: ").
.PARAMETER escape
    The character to use to escape quoted text characters (default: ").
.PARAMETER connectionTimeout
    The connection timeout to use when executing the query (default: 0; which runs until the query completes).
.PARAMETER encoding
    The encoding to use for the exported file (default: utf-8).
.PARAMETER dateFormat
    The .NET formatting mask to use when rendering DateTime objects (default: "yyyy-MM-dd HH:mm:ss").
.INPUTS
  None
.OUTPUTS
  None
.NOTES
  Version:        1.0
  Author:         Dan G. Switzer, II
  Creation Date:  April 21, 2021
  Purpose/Change: Initial script development
  
.EXAMPLE
  .\query_to_csv.ps1 -database "my-database-name" -output ".\exported-results.csv" -query "select * from tableName"
#>
# declare the parameters that can be used
param (
	  [Parameter(Mandatory, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)][string]$query
	, [Parameter(Mandatory)][string]$output
	, [string]$database=""
	, [string]$datasource="localhost,1433"
	, [string]$dbconnection=""
	, [string]$delim=","
	, [string]$newline="`r`n"
	, [string]$quote='"'
	, [string]$escape='"'
	, [string]$connectionTimeout=0
	, [string]$encoding="utf-8"
	# https://docs.microsoft.com/en-us/dotnet/standard/base-types/custom-date-and-time-format-strings
	, [string]$dateFormat="yyyy-MM-dd HH:mm:ss"
	#, [switch]$verbose
)

# make sure to resolve paths to the current working folder
$outputFile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($output);

if( $dbconnection -ne "" ){
	$connectionString = $dbconnection;
} elseif( $database -ne "" ){
	$connectionString = "Data Source=$datasource; Database=$database; Trusted_Connection=True;";
} else {
	Write-Error -Message "You must supply either the -database or -dbconnection parameters." -Category NotSpecified -ErrorAction Stop
}

# we want to track the execution time of our export process
$startTime = $(get-date);

# Powershell functions are very slow, so we use C# style classes to improve performance
#
# However, even these are considerably slower than inline object, so we use inline logic
# in our maing loop.
class Helpers { 
	static [string] CsvFormat([string]$value, [string]$quoteChar, [string]$escapeChar){
		# Rules for quoting strings from RFC4180 2.6 (https://tools.ietf.org/html/rfc4180)
		if( $value -match '\s|,|"' ){
			return $quoteChar + $value.Replace('"', $escapeChar + '"') + $quoteChar;
		} else {
			return $value;
		}
	}
}

# program defaults
$fileEncoding = $encoding;
$columnDelimiter = $delim;
$rowDelimiter = $newline;
$quoteChar = $quote;
$escapeChar = $escape;
# the more frequently we update the screen, the slower the script will run since writing to the host is an expensive operation
$exportUpdateInterval = 1000;

Write-Output "Starting the export... ";

try {
	# we need to create a new file (not append) for writing our CSV stream
	$csvStreamWriter = New-Object System.IO.StreamWriter $outputFile, $false, ([System.Text.Encoding]::GetEncoding($fileEncoding));
	# set the row delimiter to use
	$csvStreamWriter.NewLine = $rowDelimiter;

	$sqlConn = New-Object System.Data.SqlClient.SqlConnection $connectionString;
	$sqlCmd = New-Object System.Data.SqlClient.SqlCommand;
	$sqlCmd.Connection = $sqlConn;
	$sqlCmd.CommandTimeout = $connectionTimeout;
	$sqlCmd.CommandText = $query;
	$sqlConn.Open();
	# we want to open recordset and stream the results so we do not load all the records into memory
	$reader = $sqlCmd.ExecuteReader();

	Write-Output ("Query initialized in " + ("{0:HH:mm:ss}" -f ([datetime]($(get-date) - $startTime).Ticks)) + "...");

	# we do not want header output to show up in the stdout so it's not piped to output when we save it
	Write-Host "Getting headers...";
	
	# initialize the column headers
	$columns = @()
	for ( $columnCount=0 ; $columnCount -lt $reader.FieldCount; $columnCount++ ){
		$columns += @($columnCount);
		$icon = if( $columnCount -lt $reader.FieldCount-1 ){ [char]0x251c } else { [char]0x2514 };
		if( $VerbosePreference -eq "Continue" ){
			# we do not want header output to show up in the stdout so it's not piped to output when we save it
			Write-Host (" {0} {1} | {2} [{3}]" -f $icon, $reader.GetName($columnCount), $reader.GetDataTypeName($columnCount), $reader.GetFieldType($columnCount).Name);
		}
	}
	if( $VerbosePreference -eq "Continue" ){
		# print a blank line after generating the headers
		Write-Host "";
	}

	# Write Header
	$csvStreamWriter.Write([Helpers]::CsvFormat($reader.GetName(0).ToString().Trim(), $quoteChar, $escapeChar));
	for( $idx=1; $idx -lt $reader.FieldCount; $idx++ ){ 
		$csvStreamWriter.Write($($columnDelimiter + ([Helpers]::CsvFormat($reader.GetName($idx).ToString().Trim(), $quoteChar, $escapeChar))));
	}

	# Close the header line
	$csvStreamWriter.WriteLine(""); 

	Write-Output ("Exported header with {0:n0} columns" -f $reader.FieldCount);

	$rowsExported = 0;
	while( $reader.Read() ){
		$rowsExported++;

		for( $idx=0; $idx -lt $columns.Length; $idx++ ){
			# functions in Powershell are ***extremely*** slow, we avoid and just put our code inline
			$value = $reader.GetValue($idx);
			$type = $value.GetType().Name;

#			Write-Host ("{0} {1}" -f $value, $type);

			# for Date/Time operations, we should re-cast as a string
			if( ($type -eq "DateTime") ){
				$value = $value.ToString($dateFormat);
				$type = "String";
			}
			
			if( $type -eq "Boolean" ){
				$value = if( $value ){ "1" } else { "0" }
			# Rules for quoting strings from RFC4180 2.6 (https://tools.ietf.org/html/rfc4180)
			} elseif ( ($type -eq "String") -and ($value -match '[\r\n,"]') ){
				$value = $quoteChar + $value.Replace('"', $escapeChar + '"') + $quoteChar;
			}

			$columns[$idx] = $value;
		}

		# create the row data
		$rowData = [string]::Join($columnDelimiter, $columns);

		# save the row to the CSV file
		$csvStreamWriter.WriteLine($rowData);
		
		# check to see if we should update the user on the progress
		if( ($rowsExported % $exportUpdateInterval) -eq 0 ){
			# update the screen with the progress
			Write-Host -NoNewLine ("`rExported {0:n0} rows... (Runtime: {1:HH:mm:ss})" -f $rowsExported, ([datetime]($(get-date) - $startTime).Ticks));
		}
	}
	
	# since we did some updating on the screen, we need to clear it
	if( $rowsExported -gt $exportUpdateInterval ){
		# we add some spaces to make sure the entire line is overwitten
		Write-Host -NoNewLine "`r                                                                `r";
	}

	Write-Output ("Exported {0:n0} row(s)" -f $rowsExported);
	Write-Output ("Exported results to ""$outputFile""");
	Write-Output ("Finished in " + "{0:HH:mm:ss}" -f ([datetime]($(get-date) - $startTime).Ticks) + "!");
} catch {
	# the $_ outputs the actual exception message
	Write-Error "`n `nAn error occurred has occurred executing the script:`n `n$_`n "
} finally {
	# make sure to close all the streams
	if( $reader ){
		# If the script was aborted (such as via CTRL+C, we need to try and cancel the query or it will run to completion
		# and this could end up taking a long time if returning thousands of rows. By attempting to cancel the query, we can
		# halt execution immediately and the script should just stop.
		#
		# NOTE - Calling cancel on a completed query will do nothing. 
		$sqlCmd.Cancel();
		
		$reader.Close();
	}
	if( $sqlConn ){
		$sqlConn.Close();
	}
	if( $csvStreamWriter ){
		$csvStreamWriter.Close();
	}
}
