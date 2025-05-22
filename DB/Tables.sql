CREATE TABLE FACTURAS
(
	FACTURA INT IDENTITY(1,1) PRIMARY KEY,
	TIPO    CHAR(1)
)

--INSERT INTO FACTURAS SELECT 'A'

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
	precio MONEY
)
--select * from Impuestos
--INSERT INTO IMPUESTOS SELECT 'ITC',200,'20250101'
CREATE TABLE Impuestos
(
	id INT PRIMARY KEY IDENTITY(1,1),
	tipo VARCHAR(25),
	monto MONEY,
	fechaOperacion DATE
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