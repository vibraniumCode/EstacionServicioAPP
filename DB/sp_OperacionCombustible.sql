IF OBJECT_ID('sp_OperacionCombustible') IS NOT NULL
    DROP PROCEDURE sp_OperacionCombustible
GO
CREATE PROCEDURE sp_OperacionCombustible
(
	@Accion NVARCHAR(10),          -- 'INSERTAR', 'MODIFICAR', 'ELIMINAR'
    @IdCombustible INT = NULL,         -- Necesario para MODIFICAR y ELIMINAR
    @Tipo NVARCHAR(100) = NULL,
	@Precio MONEY = NULL
)

AS
BEGIN
    SET NOCOUNT ON;

    IF @Accion = 'INSERTAR'
    BEGIN
        IF @Tipo IS NULL OR @Precio IS NULL
        BEGIN
            RAISERROR('Tipo y Precio son obligatorios para insertar.', 16, 1);
            RETURN;
        END

        INSERT INTO Combustible (tipo, precio)
        VALUES (@Tipo, @Precio);

		SELECT 'Nuevo tipo de combustible registrado' AS Msg
    END

    ELSE IF @Accion = 'MODIFICAR'
    BEGIN
        IF @IdCombustible IS NULL
        BEGIN
            RAISERROR('Debe especificar IdCombustible para modificar.', 16, 1);
            RETURN;
        END

        UPDATE Combustible
        SET tipo = @Tipo,
			precio = @Precio
        WHERE id = @IdCombustible;

		SELECT 'Combustible actualizado' AS Msg
    END

    ELSE IF @Accion = 'ELIMINAR'
    BEGIN
        IF @IdCombustible IS NULL
        BEGIN
            RAISERROR('Debe especificar IdCombustible para eliminar.', 16, 1);
            RETURN;
        END

        DELETE FROM Combustible
        WHERE id = @IdCombustible;

		SELECT CONCAT('Combustible ', @IdCombustible,' eliminado') AS Msg
    END

    ELSE
    BEGIN
        RAISERROR('Acción inválida. Use: INSERTAR, MODIFICAR o ELIMINAR.', 16, 1);
    END
END;
