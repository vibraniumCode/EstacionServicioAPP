IF OBJECT_ID('sp_genera_nro_factura') IS NOT NULL
    DROP PROCEDURE sp_genera_nro_factura
GO
CREATE PROCEDURE sp_genera_nro_factura
	@codFactura INT OUTPUT,
    @nroFactura INT OUTPUT
AS
BEGIN
    SET NOCOUNT ON;

	DECLARE @ultimoCod INT, @ultimoNro INT;

	SELECT TOP 1 
		@ultimoCod = codFactura,
		@ultimoNro = nroFactura
	FROM FACTURAS
	ORDER BY codFactura, nroFactura DESC

	IF @ultimoCod IS NULL OR @ultimoNro IS NULL
    BEGIN
        SET @codFactura = 1;
        SET @nroFactura = 1;
    END
    ELSE IF @ultimoNro = 99999999
    BEGIN
        SET @codFactura = @ultimoCod + 1;
        SET @nroFactura = 1;
    END
    ELSE
    BEGIN
        SET @codFactura = @ultimoCod;
        SET @nroFactura = @ultimoNro + 1;
    END

	INSERT INTO FACTURAS (codFactura, nroFactura, tipo)
	VALUES (@codFactura, @nroFactura, 'A')
END

