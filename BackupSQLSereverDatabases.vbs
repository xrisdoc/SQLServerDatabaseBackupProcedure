' CommandTypeEnum Values
const adCmdUnspecified = -1
const adCmdText = 1
const adCmdTable = 2
const adCmdStoredProc = 4
const adCmdUnknown = 8
const adCmdFile = 256
const adCmdTableDirect = 512

' DataTypeEnum Values
const adBoolean = 11
const adVarChar = 200

' ParameterDirectionEnum Values
const adParamUnknown = 0
const adParamInput = 1
const adParamOutput = 2
const adParamInputOutput = 3
const adParamReturnValue = 4

' Database details.
' Must specify the relevant user and password to connect to the database.
' This user MUST have permision to be able to backup databases and be able to execute the relevant Stored Procedure.
const dbUser = "BackupOperator"
const dbPass = "aSecretP4ssw0rd"
const dbDataSource = "(local)\SQLExpress"

sub doDatabaseBackup(dbName, fileName, location, backupName, mediaName, appendDateTime)
	' Set the relevant variables.
	Dim dbCatalog : dbCatalog = dbName
	Dim dbConnectionString : dbConnectionString = "Provider=SQLOLEDB;Data Source=" & dbDataSource & ";Initial Catalog=" & dbCatalog & ";Persist Security Info=True;User ID=" & dbUser & ";Password=" & dbPass & ";"
	
	' Connect to the database and execute the relevant Stored Procedure to BACKUP the relevant database.
	' There are a few requirements in order to perform this operation.
	' These are all related to the user that is used to connect to the database. 
	' These are outlined below.
	'  - The user requires the permision to be able to backup databases (can be provided by the db_backupoperator role)
	'  - The user requires the permision to be able to execute the relevant Stored Procedure.

	' Create and open a connection to the database.
	Dim objConnection : Set objConnection = CreateObject("ADODB.Connection")
	objConnection.ConnectionString = dbConnectionString
	objConnection.Open()
	
	' Specify the Stored Procedure to call.
	' This stored porcedure MUST be set up for the current database within SQL Server.
	Dim dbBackupSP : dbBackupSP = "BackupCurrentDatabase"
	
	' Create a command that will execute the relevant stored procedure and obtain the @FullBackupLocation OUTPUT parameter.
	Dim objCommand : Set objCommand = CreateObject("ADODB.command")
	objCommand.ActiveConnection = objConnection
	objCommand.CommandText = dbBackupSP
	objCommand.CommandType = adCmdStoredProc
	objCommand.Parameters.Append(objCommand.CreateParameter("@BackupFileName", adVarChar, adParamInput, 100, fileName))
	objCommand.Parameters.Append(objCommand.CreateParameter("@BackupLocation", adVarChar, adParamInput, 150, location))
	objCommand.Parameters.Append(objCommand.CreateParameter("@BackupName", adVarChar, adParamInput, 100, backupName))
	objCommand.Parameters.Append(objCommand.CreateParameter("@BackupMediaName", adVarChar, adParamInput, 100, mediaName))
	objCommand.Parameters.Append(objCommand.CreateParameter("@AppendDateTime", adBoolean, adParamInput, 150, appendDateTime))
	objCommand.Parameters.Append(objCommand.CreateParameter("@FullBackupLocation", adVarChar, adParamOutput, 250))
	objCommand.Execute
	
	' Obtain the value from the FullBackupLocation OUTPUT parameter.
	Dim fullBackupLocation : fullBackupLocation = objCommand.Parameters("@FullBackupLocation").Value
	
	' Terminate the command.
	Set objCommand = Nothing
	
	' Terminate the connection to the database.
	objConnection.Close
	Set objConnection = Nothing
	
	' Output a message indicating that the database has been backed up.
	WScript.Echo "SUCCESS: The database " & dbCatalog & " on " & dbDataSource & " was backed up to " & fullBackupLocation & ""
end sub

sub Main()
	' Declare the varaibles that will hold the values obtaind from the parameters passed to this script
	' These variables will hold the details about what database to backup and additional customisations.
	Dim dbName : dbName = null
	Dim fileName : fileName = null
	Dim location : location = null
	Dim backupName : backupName = null
	Dim mediaName : mediaName = null
	Dim appendDateTime : appendDateTime = null
	
	' Set the relevant variables to customise the database backup.
	If NOT IsEmpty(WScript.Arguments.Named("dbName")) Then dbName = WScript.Arguments.Named("dbName")
	If NOT IsEmpty(WScript.Arguments.Named("fileName")) Then fileName = WScript.Arguments.Named("fileName")
	If NOT IsEmpty(WScript.Arguments.Named("location")) Then location = WScript.Arguments.Named("location")
	If NOT IsEmpty(WScript.Arguments.Named("backupName")) Then backupName = WScript.Arguments.Named("backupName")
	If NOT IsEmpty(WScript.Arguments.Named("mediaName")) Then mediaName = WScript.Arguments.Named("mediaName")
	If NOT IsEmpty(WScript.Arguments.Named("appendDateTime")) AND Lcase(WScript.Arguments.Named("appendDateTime")) = "true" Then appendDateTime = true
	
	' A check must be made to see if the database name has been specifed.
	' We need to know what database to backup before we attempt to back it up.
	If IsNull(dbName) = false AND Len(dbName) > 0 Then
		' Perform the database backup.
		doDatabaseBackup dbName, fileName, location, backupName, mediaName, appendDateTime
	Else
		' A valid database was not specified.
		WScript.Echo "ERROR: No valid database specified"
	End If
end sub

' Execute the script.
Main()