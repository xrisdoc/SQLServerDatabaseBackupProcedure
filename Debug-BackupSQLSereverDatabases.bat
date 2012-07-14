:: This batch file if for debuging purposes.
:: There is a pause command at the end of the batch file and this is to allow developers to see the output from the batch file that is called.
:: If the pause command was not included, then the Command window will close automatically.

:: Batch files can be executed by just referencing the file name, but this does not give control back to the calling batch file and therefore the calling batch file will be incomplete.
:: Using the CALL statement, the batch file will return to the calling batch file after it completes.

:: Call the batch file to backup all of the relevant SQL Databases
call BackupSQLSereverDatabases.bat

pause