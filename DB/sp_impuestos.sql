/*
SELECT * FROM IMPUESTOS
EXEC sp_impuestos 'ITC', 1700, '20250616'
SELECT * FROM IMPUESTOS
*/


CREATE OR ALTER PROCEDURE sp_impuestos
(
	@impuesto VARCHAR(25),
	@monto MONEY,
	@fechaOperacion DATE
)

AS
BEGIN
    SET NOCOUNT ON;
	
	DECLARE @AÑO INT, @MES INT
	DECLARE @mensaje NVARCHAR(100)

	SELECT TOP 1
		@AÑO = YEAR(fechaOperacion), 
		@MES = MONTH(fechaOperacion)
	FROM Impuestos ORDER BY id DESC
	
	IF @AÑO = YEAR(@fechaOperacion) AND @MES = MONTH(@fechaOperacion)
		SELECT CONCAT('Ya se encuentra cargado el impuesto ITC para el mes ', FORMAT(@fechaOperacion, 'MM/yyyy')) AS MENSAJE
		RETURN

	INSERT INTO Impuestos(tipo, monto, fechaOperacion)
        VALUES (@impuesto, @monto, @fechaOperacion);
END