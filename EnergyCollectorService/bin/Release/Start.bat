set current_dir=%~dp0
sqlcmd -S localhost\SQLExpress -U sa -P sa -i %current_dir%CreateEnergy.sql
sc create EnergyCollectorService binPath= "%current_dir%EnergyCollectorService.exe" start= "auto"
sc start EnergyCollectorService
pause