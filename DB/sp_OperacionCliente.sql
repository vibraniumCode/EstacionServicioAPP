/*
-- Insertar cliente
EXEC sp_OperacionCliente 
    @Accion = 'INSERTAR',
    @Nombre = 'Empresa ABC',
	@Direccion = 'dIRECCION TEST',
    @CUIT = '30-12345678-9';

-- Modificar cliente
EXEC sp_OperacionCliente 
    @Accion = 'MODIFICAR',
    @IdCliente = 1,
    @Nombre = 'Empresa ABC S.A.',
    @CUIT = '30-12345678-9';

-- Eliminar cliente
EXEC sp_OperacionCliente 
    @Accion = 'ELIMINAR',
    @IdCliente = 1;

*/

CREATE OR ALTER PROCEDURE sp_OperacionCliente
(
	@Accion NVARCHAR(10),          -- 'INSERTAR', 'MODIFICAR', 'ELIMINAR'
    @IdCliente INT = NULL,         -- Necesario para MODIFICAR y ELIMINAR
    @Nombre NVARCHAR(100) = NULL,
	@Direccion NVARCHAR(255) = NULL,
    @CUIT NVARCHAR(20) = NULL
)

AS
BEGIN
    SET NOCOUNT ON;

    IF @Accion = 'INSERTAR'
    BEGIN
        IF @Nombre IS NULL OR @CUIT IS NULL
        BEGIN
            RAISERROR('Nombre y CUIT son obligatorios para insertar.', 16, 1);
            RETURN;
        END

        INSERT INTO Clientes (Nombre, Direccion, CUIT)
        VALUES (@Nombre, @Direccion, @CUIT);

		SELECT 'Nuevo cliente registrado' AS Msg
        --SELECT SCOPE_IDENTITY() AS NuevoIdCliente;
    END

    ELSE IF @Accion = 'MODIFICAR'
    BEGIN
        IF @IdCliente IS NULL
        BEGIN
            RAISERROR('Debe especificar IdCliente para modificar.', 16, 1);
            RETURN;
        END

        UPDATE Clientes
        SET Nombre = @Nombre,
			Direccion = @Direccion,
            CUIT = @CUIT
        WHERE id = @IdCliente;

		SELECT 'Cliente actualizado' AS Msg
    END

    ELSE IF @Accion = 'ELIMINAR'
    BEGIN
        IF @IdCliente IS NULL
        BEGIN
            RAISERROR('Debe especificar IdCliente para eliminar.', 16, 1);
            RETURN;
        END

        DELETE FROM Clientes
        WHERE id = @IdCliente;

		SELECT CONCAT('Cliente ', @IdCliente,' eliminado') AS Msg
    END

    ELSE
    BEGIN
        RAISERROR('Acción inválida. Use: INSERTAR, MODIFICAR o ELIMINAR.', 16, 1);
    END
END;
