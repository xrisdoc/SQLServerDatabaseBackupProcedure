::CSCRIPT BackupSQLSereverDatabases.vbs /dbName:"" /fileName:"" /location:"" /backupName:"" /mediaName:"" /appendDateTime:""
::CSCRIPT BackupSQLSereverDatabases.vbs /dbName:"dbDevelopment" /appendDateTime:"True"

CSCRIPT BackupSQLSereverDatabases.vbs /dbName:"dbDevelopment" /location:"C:\temp"
CSCRIPT BackupSQLSereverDatabases.vbs /dbName:"dbMovies" /location:"C:\temp"
CSCRIPT BackupSQLSereverDatabases.vbs /dbName:"dbTest" /location:"C:\temp"
:: Copy/paste the line above for each database that is to be backed up using this procedure.