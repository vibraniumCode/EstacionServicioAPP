--sp_comprobante 0001,00000124,5

IF OBJECT_ID('sp_comprobante') IS NOT NULL
    DROP PROCEDURE sp_comprobante
GO
CREATE PROCEDURE sp_comprobante
(
	@codigoFactura INT,
	@nroFactura INT,
	@empresa INT
)

AS
BEGIN
	SET NOCOUNT ON;

	--IF @empresa = 1 
		BEGIN
			WITH UltimosImpuestos AS (
				SELECT 
					ei.*,
					ROW_NUMBER() OVER (PARTITION BY ei.idEmpresa, ei.idTipo ORDER BY ei.fechaOperacion DESC) AS rn
				FROM Empresa_Impuesto ei 
				WHERE ei.idEmpresa = @empresa 
			)

			SELECT 
				f.codFactura,
				f.nroFactura,
				e.nombre,
				e.cuit,
				e.iibb,
				e.domicilio,
				e.localidad,
				e.contacto, 
				e.inicio, 
				FORMAT(f.fecEmision, 'dd/MM/yyyy') AS FechaEmision, 
				CONVERT(time(0), f.horaEmision) AS HoraEmision, 
				c.nombre, 
				c.Cuit, 
				c.Direccion, 
				cb.id, 
				cb.tipo, 
				tc.litros, 
				cb.precio, 
				CAST(fi.imp_total / NULLIF(tc.litros,0) AS decimal(18,6)) AS PrecioPorLitro,
				CAST(tc.monto_s AS decimal(18,2)) AS neto,
				impuestos.ITC,
				impuestos.IDC,
				impuestos.IINI,
				fi.imp_iva,
				--t.tipo,
				--ui.monto AS MontoImpuesto,
				fi.impuesto_total,
				CAST(fi.imp_total AS decimal(18,2)) AS TOTAL,
				fi.imp_Hidr_Carb,
				fi.imp_Comb_Liq,
				fi.imp_Mat_Var
			FROM Facturacion f 
			INNER JOIN Facturacion_importe fi 
				ON fi.codFactura = f.codFactura AND fi.nroFactura = f.nroFactura 
			INNER JOIN Empresas e 
				ON e.id = fi.empresa 
			INNER JOIN Facturacion_dato fd 
				ON fd.codFactura = f.codFactura AND fd.nroFactura = f.nroFactura 
			INNER JOIN clientes c 
				ON c.id = fd.idCliente 
			INNER JOIN tcarga_combustible tc 
				ON tc.codFactura = f.codFactura AND tc.nroFactura = f.nroFactura 
			INNER JOIN Combustible cb 
				ON cb.id = tc.idTipo 
			--INNER JOIN UltimosImpuestos ui 
			--	ON ui.idEmpresa = e.id AND ui.rn = 1
			--INNER JOIN Timpuestos t
			--	ON t.id = ui.idTipo
			OUTER APPLY (
				SELECT
					MAX(CASE WHEN ei.idTipo = 1 THEN ei.monto END) AS ITC,
					MAX(CASE WHEN ei.idTipo = 2 THEN ei.monto END) AS IDC,
					MAX(CASE WHEN ei.idTipo = 3 THEN ei.monto END) AS IINI,
					MAX(CASE WHEN ei.idTipo = 4 THEN ei.monto END) AS IVA
				FROM (
					SELECT *,
						ROW_NUMBER() OVER (PARTITION BY idTipo ORDER BY fechaOperacion DESC) AS rn
					FROM Empresa_Impuesto
					WHERE idEmpresa = e.id
				) ei
				WHERE ei.rn = 1
			) AS impuestos

			WHERE 
				f.codFactura = @codigoFactura 
				AND f.nroFactura = @nroFactura
		END
END

--SELECT * FROM Empresa_Impuesto where idEmpresa = 1
--SELECT * FROM Facturacion_importe where nroFactura = 124