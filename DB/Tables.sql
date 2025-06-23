CREATE TABLE FACTURAS
(
	codFactura INT NOT NULL,
    nroFactura INT NOT NULL,
    tipo CHAR(1) NOT NULL,
	PRIMARY KEY (codFactura, nroFactura)
)

--INSERT INTO FACTURAS SELECT 0,0,'A'

CREATE TABLE clientes 
(
	id INT PRIMARY KEY IDENTITY(1,1),
	nombre NVARCHAR(100),
	direccion NVARCHAR(255),
	cuit NVARCHAR(20)
)

CREATE TABLE Combustible
(
	id INT PRIMARY KEY IDENTITY(1,1),
	tipo VARCHAR(25),
	precio NUMERIC(18,4)
)

CREATE TABLE Mes
(
	id INT IDENTITY(1,1),
	mes VARCHAR(10)
)

INSERT INTO Mes (mes) VALUES
('Enero'), 
('Febrero'), 
('Marzo'), 
('Abril'), 
('Mayo'), 
('Junio'), 
('Julio'), 
('Agosto'), 
('Septiembre'), 
('Octubre'), 
('Noviembre'), 
('Diciembre')



--drop table tcarga_combustible
CREATE TABLE tcarga_combustible
(
	id INT PRIMARY KEY IDENTITY(1,1),
	idTipo INT NOT NULL,
	litros NUMERIC(9,4) NOT NULL,
	monto_s NUMERIC(18,4) NOT NULL,
	codFactura INT NOT NULL,
	nroFactura INT NOT NULL
)

INSERT INTO tcarga_combustible (idTipo, litros, monto_s,codFactura,nroFactura) 
VALUES (4,10,80000,1,1)

SELECT * FROM tcarga_combustible

CREATE TABLE Facturacion
(
	codFactura INT NOT NULL,
	nroFactura INT NOT NULL,
	fecEmision DATE NOT NULL,
	horaEmision TIME NOT NULL
)

INSERT INTO Facturacion (codFactura, nroFactura, fecEmision, horaEmision)
VALUES (00001,00000001,CAST(GETDATE() AS DATE), CAST(GETDATE() AS TIME))

CREATE TABLE Facturacion_dato
(
	idDato INT IDENTITY(1,1),
	codFactura INT NOT NULL,
	nroFactura INT NOT NULL,
	idCliente INT NOT NULL
)

SELECT * FROM Facturacion_dato

--INSERT INTO Facturacion_dato (codFactura, nroFactura, idCliente)
--VALUES (00001,00000001,1)

CREATE TABLE Facturacion_importe
(
	idImporte INT IDENTITY(1,1),
	codFactura INT NOT NULL,
	nroFactura INT NOT NULL,
	imp_neto NUMERIC(18,4) NOT NULL,
	imp_iva NUMERIC(18,4) NOT NULL,
	imp_itc NUMERIC(18,4),
	imp_idc NUMERIC(18,4),
	imp_internos NUMERIC(18,4),
	impuesto_total NUMERIC(18,4) NOT NULL,
	imp_Hidr_Carb NUMERIC(18,4),
	imp_Comb_Liq NUMERIC(18,4),
	imp_Mat_Var NUMERIC(18,4),
	imp_total NUMERIC(18,4) NOT NULL,
	empresa INT NOT NULL
)

--INSERT INTO Facturacion_importe (codFactura, nroFactura, imp_neto, imp_iva, imp_itc, imp_total)
--VALUES (00001,00000001,100,10,10000,10000)

--drop table Impuestos
--drop table Timpuestos

CREATE TABLE Empresas
(
    id INT PRIMARY KEY IDENTITY(1,1),
    nombre NVARCHAR(100) NOT NULL,
    cuit NVARCHAR(15) NOT NULL,
	iibb VARCHAR(15) NOT NULL,
    domicilio NVARCHAR(255),
    localidad NVARCHAR(100),
    contacto NVARCHAR(50),
	inicio DATE
)

CREATE TABLE Timpuestos (
	id INT PRIMARY KEY IDENTITY(1,1),
	tipo VARCHAR(50) NOT NULL
);

CREATE TABLE Empresa_Impuesto (
    id INT PRIMARY KEY IDENTITY(1,1),
    idEmpresa INT NOT NULL,
    idTipo INT NOT NULL,
    monto NUMERIC(9,4) NOT NULL,
    fechaOperacion DATE NOT NULL,

    FOREIGN KEY (idEmpresa) REFERENCES Empresas(id),
    FOREIGN KEY (idTipo) REFERENCES Timpuestos(id)
);

CREATE TABLE Cierre
(
	id INT IDENTITY(1,1),
	codFactura INT,
	nroFactura INT,
	total NUMERIC(18,4)
)
/*
INSERT INTO Timpuestos SELECT 'ITC'
INSERT INTO Timpuestos SELECT 'IDC'
INSERT INTO Timpuestos SELECT 'IMPUESTO ITERNO A NIVEL ITEM'
INSERT INTO Timpuestos SELECT 'IVA'
INSERT INTO Timpuestos SELECT 'Imp. Hidr. Carb.'
INSERT INTO Timpuestos SELECT 'Imp. Comb. Liq.'
INSERT INTO Timpuestos SELECT 'Imp. Matanza Variable 1.5%'
*/

/*
insert into Empresa_Impuesto select 1,1,12.1850,getdate()
insert into Empresa_Impuesto select 1,2,962.02,getdate()
insert into Empresa_Impuesto select 1,3,43782.09,getdate()
insert into Empresa_Impuesto select 1,4,0.21,getdate()

insert into Empresa_Impuesto select 2,1,12.185,getdate()
insert into Empresa_Impuesto select 2,3,16970.68,getdate()
insert into Empresa_Impuesto select 2,4,0.21,getdate()

insert into Empresa_Impuesto select 3,1,15.698,getdate()
insert into Empresa_Impuesto select 3,3,29818.14,getdate()
insert into Empresa_Impuesto select 3,4,0.21,getdate()

insert into Empresa_Impuesto select 4,1,12.1850,getdate()
insert into Empresa_Impuesto select 4,2,962.02,getdate()
insert into Empresa_Impuesto select 4,3,43782.09,getdate()
insert into Empresa_Impuesto select 4,4,0.21,getdate()

insert into Empresa_Impuesto select 5,1,1977.62,getdate()
insert into Empresa_Impuesto select 5,4,0.21,getdate()
insert into Empresa_Impuesto select 5,5,635.17,getdate()
insert into Empresa_Impuesto select 5,6,10369.00,getdate()
insert into Empresa_Impuesto select 5,7,783.63,getdate()
*/





/*
INSERT INTO Empresas (nombre, cuit, iibb, domicilio, localidad, contacto, inicio) 
VALUES ('PETRORAFAELA S.R.L','33-68439985-9','0410219915','RUTA 34 ESQ- CHACABUCO CP: 2300','RAFAELA - SANTA FE','03492-426826','19970506')

INSERT INTO Empresas (nombre, cuit, iibb, domicilio, localidad, contacto, inicio) 
VALUES ('VALCARA SA','33-71063126-9','333-71063126-9','RUTA 3 Y RUTA 205 CAÑUELAS (1814)','PROVINCIA DE BUENOS AIRES','','20131101')

INSERT INTO Empresas (nombre, cuit, iibb, domicilio, localidad, contacto, inicio) 
VALUES ('EST. DE SERVICIO YPF','30-70963914-1','30-70963914-1','RUTA 205 KM 138,5 - C.P.7295','ROQUE PEREZ - PCIA DE BUENOS AIRES','','20060424')

INSERT INTO Empresas (nombre, cuit, iibb, domicilio, localidad, contacto, inicio) 
VALUES ('GNC GUERNICA SOCIEDAD DE RESPONSABILIDAD','30-71591180-5','30715911805','','','','20180201')

INSERT INTO Empresas (nombre, cuit, iibb, domicilio, localidad, contacto, inicio) 
VALUES ('OPERADORA DE ESTACIONES DE SERVICIO S.A.','30-67677449-5','901-163472-1',' AU TTE GRAL PABLO RICCHIERI KM 36.250 (Desc) - La Matanza PBA','Machado Nuevas 515 - CABA (1080AKX)','','19991109')
*/