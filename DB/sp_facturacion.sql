/*
exec sp_facturacion 2, 5, 1, 304.25,'20250430',1
exec sp_facturacion 2,5,1,304.25,'20250430',0


select * from combustible
select * from empresas
select * from Empresa_Impuesto ei INNER JOIN 
				Timpuestos ti on ti.id = ei.idTipo where ei.idEmpresa = 3
exec sp_facturacion 1, 5, 1, 10,'20250605',1

exec sp_facturacion 2, 5, 1, 304.25,'20250430',1

EXEC sp_facturacion 1,1,3,415.0200,'20250601',1
EXEC sp_facturacion 2,1,1,304.25,'20250601',1
EXEC sp_facturacion 3,1,1,331.4800,'20250601',1
EXEC sp_facturacion 5,1,4,52.3750,'20250601',1
*/
IF OBJECT_ID('sp_facturacion') IS NOT NULL
    DROP PROCEDURE sp_facturacion
GO
CREATE PROCEDURE sp_facturacion
(
	@empresa INT = NULL,
	@idCliente INT = NULL,
	@idCarga INT = NULL,
	@litros NUMERIC(9,4) = NULL,
	@fechaEmision DATE = NULL,
	@control BIT
)

AS
BEGIN
    SET NOCOUNT ON;

	DECLARE 
		@horaEmision TIME,
		@cod INT, 
		@nro INT,
		@imp_litros NUMERIC(9,4),
		@neto NUMERIC(18,4),
		@neto_iva NUMERIC(18,4),
		@lt_itc NUMERIC(18,4),	
		@impInterno NUMERIC(18,4),
		@imp_idc NUMERIC(18,4),
		@imp_itc NUMERIC(18,4),
		@neto_gravado NUMERIC(18,4),
		@imp_1 NUMERIC(18,4), 
		@imp_2 NUMERIC(18,4), 
		@imp_3 NUMERIC(18,4),
		@imp_total NUMERIC(18,4),
		@imp_otr_tributo NUMERIC(18,4)
		
	SET @horaEmision  = CAST(GETDATE() AS TIME)

	IF (@control = 1) 
		BEGIN
			-- Calcular precio neto (litros × precio unitario) *
			SET @imp_litros = (SELECT PRECIO FROM Combustible WHERE id = @idCarga)
			SET @neto = @litros * @imp_litros

			-- Calcular IVA (21% del neto) *
			SELECT TOP 1
				@neto_iva = @neto * ei.monto
			FROM 
				Empresa_Impuesto ei INNER JOIN 
				Timpuestos ti on ti.id = ei.idTipo 
			WHERE 
				ti.tipo = 'IVA' AND 
				ei.idEmpresa = @empresa
			ORDER BY 
				ei.fechaOperacion DESC 

			-- Calcular impuestos interno * 
			--SELECT 
			--	@impInterno = SUM(CASE WHEN ei.idTipo NOT IN (1,4) THEN ei.monto ELSE 0 END)
			--FROM 
			--	Empresa_Impuesto ei INNER JOIN 
			--	Timpuestos ti ON ti.id = ei.idTipo 
			--WHERE 
			--	ei.idEmpresa = @empresa

			SELECT 
				@impInterno = SUM(ei.monto)
			FROM Empresa_Impuesto ei
			JOIN (
				SELECT idTipo, MAX(fechaOperacion) AS fechaMax
				FROM Empresa_Impuesto
				WHERE idEmpresa = @empresa AND idTipo NOT IN (1,4)
				GROUP BY idTipo
			) ultimos ON ei.idTipo = ultimos.idTipo AND ei.fechaOperacion = ultimos.fechaMax
			WHERE ei.idEmpresa = @empresa AND ei.idTipo NOT IN (1,4)

			SET @lt_itc = (SELECT TOP 1 monto FROM 
							Empresa_Impuesto ei JOIN Timpuestos ti ON ti.id = ei.idTipo 
							WHERE ei.idEmpresa = @empresa AND ti.tipo = 'ITC' 
							ORDER BY ei.fechaOperacion DESC)
							

			IF @empresa = 1 --PETRORAFAELA SRL
				BEGIN							
					SELECT 
						@litros                                                                                 AS [Litros cargados],
						@imp_litros                                                                             AS [Precio por litro],
						@lt_itc                                                                                 AS [Impuesto ITC por litro],
						(@neto + @neto_iva + @impInterno ) / @litros                                            AS [Impuesto detallados],
						@neto                                                                                   AS [Subtotal sin impuestos],
						@neto_iva                                                                               AS [IVA],
						(SELECT TOP 1 monto from Empresa_Impuesto where idTipo = 3 AND idEmpresa = @empresa ORDER BY fechaOperacion DESC) AS [Impuesto interno],
						(SELECT TOP 1 monto from Empresa_Impuesto where idTipo = 2 AND idEmpresa = @empresa ORDER BY fechaOperacion DESC) AS [Impuesto IDC],
						@impInterno                                                                             AS [Total de impuestos],
						(@neto + @neto_iva + @impInterno)                                                       AS [Total a pagar]
				END

				IF @empresa = 2 --VALCARA SA
					BEGIN
						SET @neto_gravado = (@neto - @impInterno) / 1.21  
                                                
						SELECT 
							@litros                                                                                 AS [Litros cargados],
							@imp_litros                                                                             AS [Precio por litro],
							@neto_gravado                                                                           AS [Neto Gravado],
							@neto_gravado * 0.21																	AS [IVA],
							@impInterno                                                                             AS [Impuesto interno],
							@lt_itc                                                                                 AS [Impuesto ITC por litro],
							@neto                                                                                   AS [Total a pagar]
					END

				IF @empresa = 3 --YPF
					BEGIN
						--SET @neto_gravado = (@neto - @impInterno) / 1.21  
						                                                         
						SELECT 
							@litros                                                                                 AS [Litros cargados],
							@imp_litros                                                                             AS [Precio por litro],
							@neto                                                                                   AS [Neto Gravado],
							@neto * 0.21																	        AS [IVA],
							@impInterno                                                                             AS [Impuesto interno],
							@lt_itc                                                                                 AS [Impuesto ITC por litro],
							@neto + (@neto * 0.21) + @impInterno                                                    AS [Total a pagar]
					END

				IF @empresa = 4 --GNC GUERNICA 
					BEGIN                                                  
						SELECT 
							@litros                                                                                 AS [Litros cargados],
							@imp_litros                                                                             AS [Precio por litro],
							@neto                                                                                   AS [Neto Gravado],
							@neto * 0.21																	        AS [IVA],
							@impInterno                                                                             AS [Impuesto interno],
							@lt_itc                                                                                 AS [Impuesto ITC por litro],
							@neto + (@neto * 0.21) + @impInterno                                                    AS [Total a pagar]
					END

				IF @empresa = 5 --OPERADORA 
					BEGIN							
						SET @imp_1 = (SELECT TOP 1 monto FROM Empresa_Impuesto ei
							WHERE ei.idEmpresa = @empresa AND ei.idTipo = 5 
							ORDER BY ei.fechaOperacion DESC)

						SET @imp_2 = (SELECT TOP 1 monto FROM Empresa_Impuesto ei
							WHERE ei.idEmpresa = @empresa AND ei.idTipo = 6 
							ORDER BY ei.fechaOperacion DESC)

						SET @imp_3 = (SELECT TOP 1 monto FROM Empresa_Impuesto ei
							WHERE ei.idEmpresa = @empresa AND ei.idTipo = 7  
							ORDER BY ei.fechaOperacion DESC)
						                                                  
						SELECT 
							@litros                                                                                 AS [Litros cargados],
							@imp_litros                                                                             AS [Precio por litro],
							@neto                                                                                   AS [Neto Gravado],
							@neto * 0.21																	        AS [IVA],
							@impInterno                                                                             AS [Impuesto interno],
							@lt_itc                                                                                 AS [Impuesto ITC por litro],
							@neto + (@neto * 0.21) + @impInterno                                                    AS [Total a pagar],
							@imp_1                                                                                  AS [Imp. Hidr. de Carbono],
							@imp_2                                                                                  AS [Imp. Combustibles Líq.],
							@imp_3                                                                                  AS [Imp. a Mat. Variables]
					END
		END
	ELSE  
		BEGIN
			EXEC sp_genera_nro_factura @codFactura = @cod OUTPUT, @nroFactura = @nro OUTPUT;

			-- Calcular precio neto (litros × precio unitario) *
			SET @imp_litros = (SELECT PRECIO FROM Combustible WHERE id = @idCarga)
			SET @neto = @litros * @imp_litros

			-- Calcular IVA (21% del neto) *
			SELECT TOP 1
				@neto_iva = @neto * ei.monto
			FROM 
				Empresa_Impuesto ei INNER JOIN 
				Timpuestos ti on ti.id = ei.idTipo 
			WHERE 
				ti.tipo = 'IVA' AND 
				ei.idEmpresa = @empresa
			ORDER BY 
				ei.fechaOperacion DESC 

			-- Calcular impuestos interno * 
			--SELECT 
			--	@impInterno = SUM(CASE WHEN ei.idTipo NOT IN (1,4) THEN ei.monto ELSE 0 END)
			--FROM 
			--	Empresa_Impuesto ei INNER JOIN 
			--	Timpuestos ti ON ti.id = ei.idTipo 
			--WHERE 
			--	ei.idEmpresa = @empresa

			SELECT 
				@imp_otr_tributo = SUM(ei.monto)
			FROM Empresa_Impuesto ei
			JOIN (
				SELECT idTipo, MAX(fechaOperacion) AS fechaMax
				FROM Empresa_Impuesto
				WHERE idEmpresa = @empresa AND idTipo NOT IN (1,4)
				GROUP BY idTipo
			) ultimos ON ei.idTipo = ultimos.idTipo AND ei.fechaOperacion = ultimos.fechaMax
			WHERE ei.idEmpresa = @empresa AND ei.idTipo NOT IN (1,4)

			SET @imp_itc = (SELECT TOP 1 monto FROM 
							Empresa_Impuesto ei JOIN Timpuestos ti ON ti.id = ei.idTipo 
							WHERE ei.idEmpresa = @empresa AND ti.tipo = 'ITC' 
							ORDER BY ei.fechaOperacion DESC)

			INSERT INTO Facturacion (codFactura, nroFactura, fecEmision, horaEmision)
					VALUES (@cod, @nro, @fechaEmision, @horaEmision)

			INSERT INTO Facturacion_dato (codFactura, nroFactura, idCliente)
					VALUES (@cod, @nro, @idCliente)

			IF @empresa = 1 --PETRORAFAELA SRL
				BEGIN		
					SET @imp_total = (@neto + @neto_iva + @imp_otr_tributo)

					INSERT INTO tcarga_combustible (idTipo, litros, monto_s, codFactura, nroFactura)
					VALUES (@idCarga, @litros, @neto, @cod, @nro);

					SET @imp_idc = (SELECT TOP 1 monto FROM Empresa_Impuesto WHERE idTipo = 2 AND idEmpresa = 1 ORDER BY fechaOperacion DESC)
					SET @impInterno = (SELECT TOP 1 monto FROM Empresa_Impuesto WHERE idTipo = 3 AND idEmpresa = 1 ORDER BY fechaOperacion DESC)

					--SET @imp_otr_tributo = @imp_idc + @impInterno

					INSERT Facturacion_importe (
						codFactura, nroFactura, imp_neto, imp_iva, imp_itc, imp_idc, imp_internos, impuesto_total, imp_Hidr_Carb, imp_Comb_Liq, imp_Mat_Var, imp_total, empresa)
					VALUES (
						@cod, @nro, @neto, @neto_iva, @imp_itc, @imp_idc, @impInterno, @imp_otr_tributo, null, null, null, @imp_total, @empresa)			
				END

				IF @empresa = 2 --VALCARA SA
					BEGIN
						SET @neto_gravado = (@neto - @imp_otr_tributo) / 1.21  
						SET @neto_iva = @neto_gravado * 0.21       
						SET @imp_total = @neto                               

						INSERT INTO tcarga_combustible (idTipo, litros, monto_s, codFactura, nroFactura)
							VALUES (@idCarga, @litros, @neto_gravado, @cod, @nro);

						INSERT Facturacion_importe (
							codFactura, nroFactura, imp_neto, imp_iva, imp_itc, imp_idc, imp_internos, impuesto_total, imp_Hidr_Carb, imp_Comb_Liq, imp_Mat_Var, imp_total, empresa)
						VALUES (
							@cod, @nro, @neto_gravado, @neto_iva, @imp_itc, null, @imp_otr_tributo, @imp_otr_tributo, null, null, null, @imp_total, @empresa)			
					END

				IF @empresa = 3 --YPF
					BEGIN
						SET @neto_iva = @neto * 0.21   
						SET @imp_total = @neto + (@neto * 0.21) + @imp_otr_tributo   
						
						INSERT INTO tcarga_combustible (idTipo, litros, monto_s, codFactura, nroFactura)
							VALUES (@idCarga, @litros, @neto, @cod, @nro);                                                     

						INSERT Facturacion_importe (
							codFactura, nroFactura, imp_neto, imp_iva, imp_itc, imp_idc, imp_internos, impuesto_total, imp_Hidr_Carb, imp_Comb_Liq, imp_Mat_Var, imp_total, empresa)
						VALUES (
							@cod, @nro, @neto, @neto_iva, @imp_itc, null, @imp_otr_tributo, @imp_otr_tributo, null, null, null, @imp_total, @empresa)	
					END

				IF @empresa = 4 --GNC GUERNICA 
					BEGIN        
						SET @neto_iva = @neto * 0.21   
						SET @imp_total = @neto + (@neto * 0.21) + @imp_otr_tributo   
						
						INSERT INTO tcarga_combustible (idTipo, litros, monto_s, codFactura, nroFactura)
							VALUES (@idCarga, @litros, @neto, @cod, @nro);                                                     

						INSERT Facturacion_importe (
							codFactura, nroFactura, imp_neto, imp_iva, imp_itc, imp_idc, imp_internos, impuesto_total, imp_Hidr_Carb, imp_Comb_Liq, imp_Mat_Var, imp_total, empresa)
						VALUES (
							@cod, @nro, @neto, @neto_iva, @imp_itc, null, @imp_otr_tributo, @imp_otr_tributo, null, null, null, @imp_total, @empresa)
					END

				IF @empresa = 5 --OPERADORA 
					BEGIN		
						SET @neto_iva = @neto * 0.21
						SET @imp_total = @neto + (@neto * 0.21) + @imp_otr_tributo 
										
						SET @imp_1 = (SELECT TOP 1 monto FROM Empresa_Impuesto ei
							WHERE ei.idEmpresa = @empresa AND ei.idTipo = 5 
							ORDER BY ei.fechaOperacion DESC)

						SET @imp_2 = (SELECT TOP 1 monto FROM Empresa_Impuesto ei
							WHERE ei.idEmpresa = @empresa AND ei.idTipo = 6 
							ORDER BY ei.fechaOperacion DESC)

						SET @imp_3 = (SELECT TOP 1 monto FROM Empresa_Impuesto ei
							WHERE ei.idEmpresa = @empresa AND ei.idTipo = 7  
							ORDER BY ei.fechaOperacion DESC)

						INSERT INTO tcarga_combustible (idTipo, litros, monto_s, codFactura, nroFactura)
							VALUES (@idCarga, @litros, @neto, @cod, @nro);   

						INSERT Facturacion_importe (
							codFactura, nroFactura, imp_neto, imp_iva, imp_itc, imp_idc, imp_internos, impuesto_total, imp_Hidr_Carb, imp_Comb_Liq, imp_Mat_Var, imp_total, empresa)
						VALUES (
							@cod, @nro, @neto, @neto_iva, @imp_itc, null, @imp_otr_tributo, @imp_otr_tributo, @imp_1, @imp_2, @imp_3, @imp_total, @empresa)
					END

			SELECT @cod AS codFactura, @nro AS nroFactura;
		END
END


