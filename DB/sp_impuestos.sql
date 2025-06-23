/*
SELECT * FROM Empresa_Impuesto
exec sp_impuestos 1,NULL,NULL,NULL,1,2025,NULL,'GRL'
*/

IF OBJECT_ID('sp_impuestos') IS NOT NULL
    DROP PROCEDURE sp_impuestos
GO
CREATE PROCEDURE sp_impuestos
(
	@idImpuesto INT       = NULL,
	@impuesto VARCHAR(50) = NULL,
	@monto MONEY          = NULL,
	@fechaOperacion DATE  = NULL,
	@empresa INT          = NULL,
	@anioVB INT           = NULL,
	@mesVB INT            = NULL,
	@proceso CHAR(3)      = NULL
)

AS

BEGIN
    SET NOCOUNT ON;
	
	DECLARE 
		@anio INT, 
		@mes INT
	
	IF @proceso = 'MOD'
		BEGIN
			SELECT TOP 1
				@anio = YEAR(fechaOperacion), 
				@mes = MONTH(fechaOperacion)
			FROM 
				Empresa_Impuesto 
			WHERE 
				idTipo = @idImpuesto AND 
				idEmpresa = @empresa
			ORDER BY 
				id DESC
	
			IF @anio = YEAR(@fechaOperacion) AND @mes = MONTH(@fechaOperacion)
			BEGIN
				SELECT 1, CONCAT('Ya se encuentra cargado el impuesto para el mes ', FORMAT(@fechaOperacion, 'MM/yyyy')) AS MENSAJE
				RETURN;
			END
			
			INSERT INTO Empresa_Impuesto(idEmpresa, idTipo, monto, fechaOperacion)
				VALUES (@empresa, @idImpuesto, @monto, @fechaOperacion);

			SELECT 'Monto cargado correctamente' AS Mensaje
			RETURN;
		END
	 
	IF @proceso = 'UPD'
		BEGIN
			UPDATE Empresa_Impuesto SET 
				monto = @monto,
				fechaOperacion = @fechaOperacion
			WHERE 
				idEmpresa = @empresa AND 
				idTipo = @idImpuesto

			SELECT 'Monto actualizado correctamente' AS Mensaje
			RETURN;
		END

	IF @proceso = 'GRL'
		BEGIN
			SELECT 
				ei.fechaOperacion, 
				ti.tipo, 
				ei.monto 
			FROM 
				Empresa_Impuesto ei JOIN
				Timpuestos ti ON ti.id = ei.idTipo 
			WHERE 
				YEAR(ei.fechaOperacion) = @anioVB AND
				(@mesVB = 999 OR MONTH(ei.fechaOperacion) = @mesVB) AND
				ei.idEmpresa = @empresa AND
				(@idImpuesto = 999 OR ti.id = @idImpuesto)
			ORDER BY 
				ei.fechaOperacion ASC
		END
END