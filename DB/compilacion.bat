@echo off
SET SERVER=CNP-SIST027-N
SET DATABASE=EstacionServiciosDB

sqlcmd -S %SERVER% -d %DATABASE% -E -i "sp_OperacionCliente.sql"
sqlcmd -S %SERVER% -d %DATABASE% -E -i "sp_OperacionCombustible.sql"
sqlcmd -S %SERVER% -d %DATABASE% -E -i "sp_impuestos.sql"

pause
