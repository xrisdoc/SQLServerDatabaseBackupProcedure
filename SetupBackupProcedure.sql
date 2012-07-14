USE master
GO

-- Check to see if the database login for BackupOperator exists and if it doesn't, create it.
IF NOT EXISTS(SELECT name FROM sys.sql_logins WHERE name = 'BackupOperator')
	-- [#!# - Change Required (change password accordingly) - #!#]
	-- The password used here should also be used in the VBS file that will instigate the Backup Process.
	CREATE LOGIN BackupOperator WITH PASSWORD = 'aSecretP4ssw0rd', CHECK_EXPIRATION = OFF , CHECK_POLICY = OFF
GO

-- Change to the database we are installing the BACKUP procedure on.
-- [#!# - Change Required - #!#]
USE dbTest
GO

-- Check if the BackupCurrentDatabase stored procedure already exists in the database.
-- If it does, then it will be DROPed and then re-created.
IF OBJECT_ID('BackupCurrentDatabase') IS NOT NULL
BEGIN
	DROP PROCEDURE BackupCurrentDatabase
END
GO

-- Create the BackupCurrentDatabase stored procedure.
CREATE PROCEDURE [dbo].[BackupCurrentDatabase]
	@BackupFileName NVARCHAR(100) = Null,
	@BackupLocation NVARCHAR(150) = Null,
	@BackupName NVARCHAR(100) = Null,
	@BackupMediaName NVARCHAR(100) = Null,
	@AppendDateTime BIT = 0,
	@FullBackupLocation NVARCHAR(250) OUTPUT
AS
BEGIN	
	-- Obtain the current database name.
	-- This will be used to generate the default file name for the backup.
	DECLARE @DatabaseName NVARCHAR(30)
	SET @DatabaseName = db_name()
	
	-- Set the filename for the backup file to the database name, only if the file name has not already been specified.
	IF @BackupFileName Is Null
		SET @BackupFileName = @DatabaseName;
	
	-- If the date-time has been specified to be appended to the filename.
	-- If so, then append the current date-time.
	IF @AppendDateTime = 1
	BEGIN
		DECLARE @BackupDateTime NVARCHAR(19);
		SET @BackupDateTime = CONVERT(VARCHAR(10), GETDATE(), 105) + '-' + REPLACE(CONVERT(VARCHAR(8), GETDATE(), 108), ':', '-');
		SET @BackupFileName = @BackupFileName + '(' + @BackupDateTime + ')';
	END
	
	-- Set the location of where the backup is to be located.
	-- If the backup location has not been specified, then use the default.
	IF @BackupLocation Is Null
		SET @FullBackupLocation = 'C:\SQLDatabaseBackups\' + @BackupFileName + '.bak'
	ELSE
		SET @FullBackupLocation = @BackupLocation +'\' + @BackupFileName + '.bak'
	
	-- Set the name of the backup, only if it has not already been specified.
	IF @BackupName Is Null
		SET @BackupName = 'Full Backup of ' + @DatabaseName
	
	-- Set the media name for the backup, only if it has not already been specified.
	IF @BackupMediaName Is Null
		SET @BackupMediaName = 'FileSystem_SQLServerBackups'
	
	-- Initial backup SQL query was generated using the script option on the Backup Database dialog box.
	BACKUP DATABASE @DatabaseName 
	TO DISK = @FullBackupLocation
	WITH NOFORMAT, NOINIT,
	MEDIANAME = @BackupMediaName,
	NAME = @BackupName,
	SKIP, NOREWIND, NOUNLOAD, STATS = 10;
END
GO

-- Check if the user BackupOperator for the BackupOperator login exists within the current database.
-- If not, then create it.
IF NOT EXISTS (SELECT name FROM sys.database_principals WHERE name = 'BackupOperator')
BEGIN
	CREATE USER [BackupOperator] FOR LOGIN [BackupOperator]
END

-- Add the backup operator role to the BackupOperator user on the current database.
EXEC sp_addrolemember N'db_backupoperator', N'BackupOperator'

-- Grant the execute permission to the BackupOperator user on the BackupCurrentDatabase stored procedure for the current database.
GRANT EXECUTE ON [dbo].[BackupCurrentDatabase] TO [BackupOperator]
GO