# Define parameters
$dateString = Get-Date -Format "yyyy-MM-dd"

# file name you want to keep for this report
$originalFileName = "QueuePushAudit" 				

$newFileName = "${originalFileName}_$dateString"

# server details where query would run
$serverName = "APPRDCRYSHESDB1"
$databaseName = "Kimbal_Support"

# sql query you want to run for this report
$query = "
USE [Kimbal_Support]

    SELECT TOP 1
           [Date_IST],
           [FromDateTime_IST],
           [ToDateTime_IST],
           [Data_Available_At_HES_BlockLoadProfile],
           [Data_Sent_to_Queue_BlockLoadProfile],
           [Data_Available_At_HES_DailyLoadProfile],
           [Data_Sent_to_Queue_DailyLoadProfile],
           [Data_Available_At_HES_BillingProfile],
           [Data_Sent_to_Queue_BillingProfile]
	FROM [dbo].[QueuePushAuditCountDistinct]
	order by CreatedDate_IST desc;
"	

# location you want to keep the report in your local
$outputFile = "D:\dist\$newFileName.csv"	

# Load the SQL Server module
Import-Module SqlServer

try {
    # Create a SQL connection
    $connectionString = "Server=$serverName;Database=$databaseName;User Id=KimbalHESUser;Password=3y5F@ikeTQ01;"
    $connection = New-Object System.Data.SqlClient.SqlConnection
    $connection.ConnectionString = $connectionString

    # Open the connection
    $connection.Open()

    # Create a SQL command
    $command = $connection.CreateCommand()
    $command.CommandText = $query

    # Execute the command and read the data
    $reader = $command.ExecuteReader()
    $table = New-Object System.Data.DataTable
    $table.Load($reader)

    # Close the reader and the connection
    $reader.Close()
    $connection.Close()

    # Export the data to a CSV file
    $table | Export-Csv -Path $outputFile -NoTypeInformation

    Write-Host "Data exported to $outputFile"

}
catch {
    Write-Host "An error occurred: $_"
}