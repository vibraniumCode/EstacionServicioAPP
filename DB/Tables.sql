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

CREATE TABLE Impuestos
(
	id INT PRIMARY KEY IDENTITY(1,1),
	tipo VARCHAR(25),
	monto MONEY,
	fechaOperacion DATE
)


