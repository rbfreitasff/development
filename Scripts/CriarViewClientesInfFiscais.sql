USE [Gestao]
GO

/****** Object:  View [dbo].[ClientesInfFiscais]    Script Date: 03/06/2020 13:54:48 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

DROP VIEW [dbo].[ClientesInfFiscais]
GO

CREATE VIEW [dbo].[ClientesInfFiscais]
AS 
SELECT 
	Clientes.idCliente, 
	Clientes.Nome, 
	Clientes.Nome AS Fantasia, 
	Clientes.CN AS CN, 
	Clientes.CE AS IE, 
	'.' AS Logradouro,
	Clientes.Endereco, 
	'SN' AS Numero, 
	Clientes.Bairro, 
	Municipios.Codigo AS CodigoMunicipio, 
	Municipios.Nome AS NomeMunicipio, 
	'' AS InscricaoMunicipal, 
	UFs.Codigo AS CodigoUF, 
	UFs.Sigla AS SiglaUF, 
	Clientes.CEP, 
	Clientes.Telefone_1 AS Telefone, 
	Clientes.Interestadual, 
	Clientes.ConsumidorFinal, 
	Clientes.idTipoContribuinte AS TipoContribuinte
FROM 
	Clientes,
    Municipios,
	UFs
WHERE
	Clientes.idMunicipio = Municipios.idMunicipio 
AND Municipios.UF = UFs.Sigla

GO

