sc delete EnergyCollectorService
sqlcmd -S localhost\SQLExpress -U sa -P sa -i %current_dir%DeleteEnergy.sql
pause