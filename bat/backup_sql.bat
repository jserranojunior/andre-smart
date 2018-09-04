sqlcmd -U sa -P @fgh55qdy -S SERVIDORVM01-PC\SQLEXPRESS -Q "EXEC vdlap.dbo.sp_BackupDatabases @databaseName='vdlap_novo', @backupLocation='C:\Backup\SQL\vdlap_sql_novo\', @backupType='F'"

sqlcmd -U sa -P @fgh55qdy -S SERVIDORVM01-PC\SQLEXPRESS -Q "EXEC vdlap.dbo.sp_BackupDatabases @databaseName='vdlap', @backupLocation='C:\Backup\SQL\vdlap_sql_antigo\', @backupType='F'"

 

