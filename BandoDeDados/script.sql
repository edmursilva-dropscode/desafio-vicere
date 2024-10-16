USE [BDAdministradoraCC]
GO
/****** Object:  UserDefinedFunction [dbo].[CategoriaTransacao]    Script Date: 11/10/2024 09:51:30 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO



CREATE FUNCTION [dbo].[CategoriaTransacao] (@Ativo INT)
RETURNS VARCHAR(10)
AS
BEGIN
    DECLARE @Categoria VARCHAR(10)
    IF @Ativo = 0
        SET @Categoria = 'Não'
    ELSE IF @Ativo = 1
        SET @Categoria = 'Sim'
    ELSE
        SET @Categoria = 'Não'
    RETURN @Categoria
END;
GO
/****** Object:  Table [dbo].[Transacoes]    Script Date: 11/10/2024 09:51:30 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Transacoes](
	[ID_Transacao] [int] IDENTITY(1,1) NOT NULL,
	[Numero_CPF] [varchar](14) NOT NULL,
	[ID_Corretor] [int] NOT NULL,
	[ID_Cidade] [int] NOT NULL,
	[Ativo] [int] NOT NULL,
	[Data_Transacao] [datetime] NOT NULL,
PRIMARY KEY CLUSTERED 
(
	[ID_Transacao] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[Corretores]    Script Date: 11/10/2024 09:51:30 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Corretores](
	[ID_Corretor] [int] IDENTITY(1,1) NOT NULL,
	[Nome_Corretor] [varchar](100) NOT NULL,
PRIMARY KEY CLUSTERED 
(
	[ID_Corretor] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[Clientes]    Script Date: 11/10/2024 09:51:30 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Clientes](
	[ID_Cliente] [int] IDENTITY(1,1) NOT NULL,
	[Nome_Cliente] [varchar](100) NOT NULL,
	[Numero_CPF] [varchar](14) NOT NULL,
PRIMARY KEY CLUSTERED 
(
	[ID_Cliente] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[Estados]    Script Date: 11/10/2024 09:51:30 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Estados](
	[ID_Estado] [int] IDENTITY(1,1) NOT NULL,
	[Nome] [nvarchar](100) NOT NULL,
	[Sigla] [nvarchar](2) NOT NULL,
PRIMARY KEY CLUSTERED 
(
	[ID_Estado] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Table [dbo].[Cidades]    Script Date: 11/10/2024 09:51:30 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[Cidades](
	[ID_Cidade] [int] IDENTITY(1,1) NOT NULL,
	[Nome] [nvarchar](100) NOT NULL,
	[ID_Estado] [int] NOT NULL,
PRIMARY KEY CLUSTERED 
(
	[ID_Cidade] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  View [dbo].[vw_Transacoes]    Script Date: 11/10/2024 09:51:30 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO






CREATE VIEW [dbo].[vw_Transacoes] AS
SELECT 
	t.ID_Transacao,
	c.ID_Cliente,
    c.Nome_Cliente,
    t.Numero_CPF,
	dbo.CategoriaTransacao(t.Ativo) AS Categoria,
	t.ID_Corretor,
	d.Nome_corretor,
	f.Sigla,
	e.Nome AS Nome_Cidade,
    t.Data_Transacao
FROM 
    Transacoes t
    INNER JOIN Clientes c ON t.Numero_CPF = c.Numero_CPF
	INNER JOIN Corretores d ON t.ID_Corretor = d.ID_Corretor
	INNER JOIN Cidades e ON t.ID_Cidade = e.ID_Cidade
	INNER JOIN Estados f ON e.ID_Estado = f.ID_Estado;
GO
SET IDENTITY_INSERT [dbo].[Cidades] ON 

INSERT [dbo].[Cidades] ([ID_Cidade], [Nome], [ID_Estado]) VALUES (1, N'Rio Branco', 1)
INSERT [dbo].[Cidades] ([ID_Cidade], [Nome], [ID_Estado]) VALUES (2, N'Cruzeiro do Sul', 1)
INSERT [dbo].[Cidades] ([ID_Cidade], [Nome], [ID_Estado]) VALUES (3, N'Sena Madureira', 1)
INSERT [dbo].[Cidades] ([ID_Cidade], [Nome], [ID_Estado]) VALUES (4, N'Maceió', 2)
INSERT [dbo].[Cidades] ([ID_Cidade], [Nome], [ID_Estado]) VALUES (5, N'Arapiraca', 2)
INSERT [dbo].[Cidades] ([ID_Cidade], [Nome], [ID_Estado]) VALUES (6, N'Rio Largo', 2)
INSERT [dbo].[Cidades] ([ID_Cidade], [Nome], [ID_Estado]) VALUES (7, N'Macapá', 3)
INSERT [dbo].[Cidades] ([ID_Cidade], [Nome], [ID_Estado]) VALUES (8, N'Santana', 3)
INSERT [dbo].[Cidades] ([ID_Cidade], [Nome], [ID_Estado]) VALUES (9, N'Laranjal do Jari', 3)
INSERT [dbo].[Cidades] ([ID_Cidade], [Nome], [ID_Estado]) VALUES (10, N'São Paulo', 25)
INSERT [dbo].[Cidades] ([ID_Cidade], [Nome], [ID_Estado]) VALUES (11, N'Campinas', 25)
INSERT [dbo].[Cidades] ([ID_Cidade], [Nome], [ID_Estado]) VALUES (12, N'Guarulhos', 25)
INSERT [dbo].[Cidades] ([ID_Cidade], [Nome], [ID_Estado]) VALUES (13, N'São Bernardo do Campo', 25)
INSERT [dbo].[Cidades] ([ID_Cidade], [Nome], [ID_Estado]) VALUES (14, N'Santo André', 25)
INSERT [dbo].[Cidades] ([ID_Cidade], [Nome], [ID_Estado]) VALUES (15, N'Ribeirão Preto', 25)
INSERT [dbo].[Cidades] ([ID_Cidade], [Nome], [ID_Estado]) VALUES (16, N'Osasco', 25)
INSERT [dbo].[Cidades] ([ID_Cidade], [Nome], [ID_Estado]) VALUES (17, N'Sorocaba', 25)
INSERT [dbo].[Cidades] ([ID_Cidade], [Nome], [ID_Estado]) VALUES (18, N'São José dos Campos', 25)
INSERT [dbo].[Cidades] ([ID_Cidade], [Nome], [ID_Estado]) VALUES (19, N'Santos', 25)
INSERT [dbo].[Cidades] ([ID_Cidade], [Nome], [ID_Estado]) VALUES (20, N'Mauá', 25)
INSERT [dbo].[Cidades] ([ID_Cidade], [Nome], [ID_Estado]) VALUES (21, N'São José do Rio Preto', 25)
INSERT [dbo].[Cidades] ([ID_Cidade], [Nome], [ID_Estado]) VALUES (22, N'Mogi das Cruzes', 25)
INSERT [dbo].[Cidades] ([ID_Cidade], [Nome], [ID_Estado]) VALUES (23, N'Jundiaí', 25)
INSERT [dbo].[Cidades] ([ID_Cidade], [Nome], [ID_Estado]) VALUES (24, N'Piracicaba', 25)
INSERT [dbo].[Cidades] ([ID_Cidade], [Nome], [ID_Estado]) VALUES (25, N'Rio de Janeiro', 19)
INSERT [dbo].[Cidades] ([ID_Cidade], [Nome], [ID_Estado]) VALUES (26, N'São Gonçalo', 19)
INSERT [dbo].[Cidades] ([ID_Cidade], [Nome], [ID_Estado]) VALUES (27, N'Duque de Caxias', 19)
INSERT [dbo].[Cidades] ([ID_Cidade], [Nome], [ID_Estado]) VALUES (28, N'Nova Iguaçu', 19)
INSERT [dbo].[Cidades] ([ID_Cidade], [Nome], [ID_Estado]) VALUES (29, N'Niterói', 19)
INSERT [dbo].[Cidades] ([ID_Cidade], [Nome], [ID_Estado]) VALUES (30, N'Belo Horizonte', 13)
INSERT [dbo].[Cidades] ([ID_Cidade], [Nome], [ID_Estado]) VALUES (31, N'Uberlândia', 13)
INSERT [dbo].[Cidades] ([ID_Cidade], [Nome], [ID_Estado]) VALUES (32, N'Contagem', 13)
INSERT [dbo].[Cidades] ([ID_Cidade], [Nome], [ID_Estado]) VALUES (33, N'Juiz de Fora', 13)
INSERT [dbo].[Cidades] ([ID_Cidade], [Nome], [ID_Estado]) VALUES (34, N'Betim', 13)
INSERT [dbo].[Cidades] ([ID_Cidade], [Nome], [ID_Estado]) VALUES (35, N'Porto Alegre', 21)
INSERT [dbo].[Cidades] ([ID_Cidade], [Nome], [ID_Estado]) VALUES (36, N'Caxias do Sul', 21)
INSERT [dbo].[Cidades] ([ID_Cidade], [Nome], [ID_Estado]) VALUES (37, N'Pelotas', 21)
INSERT [dbo].[Cidades] ([ID_Cidade], [Nome], [ID_Estado]) VALUES (38, N'Canoas', 21)
INSERT [dbo].[Cidades] ([ID_Cidade], [Nome], [ID_Estado]) VALUES (39, N'Santa Maria', 21)
INSERT [dbo].[Cidades] ([ID_Cidade], [Nome], [ID_Estado]) VALUES (40, N'Curitiba', 16)
INSERT [dbo].[Cidades] ([ID_Cidade], [Nome], [ID_Estado]) VALUES (41, N'Londrina', 16)
INSERT [dbo].[Cidades] ([ID_Cidade], [Nome], [ID_Estado]) VALUES (42, N'Maringá', 16)
INSERT [dbo].[Cidades] ([ID_Cidade], [Nome], [ID_Estado]) VALUES (43, N'Ponta Grossa', 16)
INSERT [dbo].[Cidades] ([ID_Cidade], [Nome], [ID_Estado]) VALUES (44, N'Cascavel', 16)
INSERT [dbo].[Cidades] ([ID_Cidade], [Nome], [ID_Estado]) VALUES (45, N'Salvador', 5)
INSERT [dbo].[Cidades] ([ID_Cidade], [Nome], [ID_Estado]) VALUES (46, N'Feira de Santana', 5)
INSERT [dbo].[Cidades] ([ID_Cidade], [Nome], [ID_Estado]) VALUES (47, N'Vitória da Conquista', 5)
INSERT [dbo].[Cidades] ([ID_Cidade], [Nome], [ID_Estado]) VALUES (48, N'Camaçari', 5)
INSERT [dbo].[Cidades] ([ID_Cidade], [Nome], [ID_Estado]) VALUES (49, N'Itabuna', 5)
INSERT [dbo].[Cidades] ([ID_Cidade], [Nome], [ID_Estado]) VALUES (50, N'Recife', 17)
INSERT [dbo].[Cidades] ([ID_Cidade], [Nome], [ID_Estado]) VALUES (51, N'Jaboatão dos Guararapes', 17)
INSERT [dbo].[Cidades] ([ID_Cidade], [Nome], [ID_Estado]) VALUES (52, N'Olinda', 17)
INSERT [dbo].[Cidades] ([ID_Cidade], [Nome], [ID_Estado]) VALUES (53, N'Caruaru', 17)
INSERT [dbo].[Cidades] ([ID_Cidade], [Nome], [ID_Estado]) VALUES (54, N'Petrolina', 17)
INSERT [dbo].[Cidades] ([ID_Cidade], [Nome], [ID_Estado]) VALUES (55, N'Fortaleza', 6)
INSERT [dbo].[Cidades] ([ID_Cidade], [Nome], [ID_Estado]) VALUES (56, N'Caucaia', 6)
INSERT [dbo].[Cidades] ([ID_Cidade], [Nome], [ID_Estado]) VALUES (57, N'Juazeiro do Norte', 6)
INSERT [dbo].[Cidades] ([ID_Cidade], [Nome], [ID_Estado]) VALUES (58, N'Maracanaú', 6)
INSERT [dbo].[Cidades] ([ID_Cidade], [Nome], [ID_Estado]) VALUES (59, N'Sobral', 6)
SET IDENTITY_INSERT [dbo].[Cidades] OFF
GO
SET IDENTITY_INSERT [dbo].[Clientes] ON 

INSERT [dbo].[Clientes] ([ID_Cliente], [Nome_Cliente], [Numero_CPF]) VALUES (1, N'Teste cliente 01', N'123.456.789-99')
INSERT [dbo].[Clientes] ([ID_Cliente], [Nome_Cliente], [Numero_CPF]) VALUES (4, N'Teste cliente 02', N'123.456.789-68')
INSERT [dbo].[Clientes] ([ID_Cliente], [Nome_Cliente], [Numero_CPF]) VALUES (5, N'Teste cliente 03', N'986.752.145-63')
SET IDENTITY_INSERT [dbo].[Clientes] OFF
GO
SET IDENTITY_INSERT [dbo].[Corretores] ON 

INSERT [dbo].[Corretores] ([ID_Corretor], [Nome_Corretor]) VALUES (1, N'Teste corretor 01')
INSERT [dbo].[Corretores] ([ID_Corretor], [Nome_Corretor]) VALUES (3, N'Teste corretor 02')
INSERT [dbo].[Corretores] ([ID_Corretor], [Nome_Corretor]) VALUES (4, N'Teste corretor 03')
SET IDENTITY_INSERT [dbo].[Corretores] OFF
GO
SET IDENTITY_INSERT [dbo].[Estados] ON 

INSERT [dbo].[Estados] ([ID_Estado], [Nome], [Sigla]) VALUES (1, N'Acre', N'AC')
INSERT [dbo].[Estados] ([ID_Estado], [Nome], [Sigla]) VALUES (2, N'Alagoas', N'AL')
INSERT [dbo].[Estados] ([ID_Estado], [Nome], [Sigla]) VALUES (3, N'Amapá', N'AP')
INSERT [dbo].[Estados] ([ID_Estado], [Nome], [Sigla]) VALUES (4, N'Amazonas', N'AM')
INSERT [dbo].[Estados] ([ID_Estado], [Nome], [Sigla]) VALUES (5, N'Bahia', N'BA')
INSERT [dbo].[Estados] ([ID_Estado], [Nome], [Sigla]) VALUES (6, N'Ceará', N'CE')
INSERT [dbo].[Estados] ([ID_Estado], [Nome], [Sigla]) VALUES (7, N'Distrito Federal', N'DF')
INSERT [dbo].[Estados] ([ID_Estado], [Nome], [Sigla]) VALUES (8, N'Espírito Santo', N'ES')
INSERT [dbo].[Estados] ([ID_Estado], [Nome], [Sigla]) VALUES (9, N'Goiás', N'GO')
INSERT [dbo].[Estados] ([ID_Estado], [Nome], [Sigla]) VALUES (10, N'Maranhão', N'MA')
INSERT [dbo].[Estados] ([ID_Estado], [Nome], [Sigla]) VALUES (11, N'Mato Grosso', N'MT')
INSERT [dbo].[Estados] ([ID_Estado], [Nome], [Sigla]) VALUES (12, N'Mato Grosso do Sul', N'MS')
INSERT [dbo].[Estados] ([ID_Estado], [Nome], [Sigla]) VALUES (13, N'Minas Gerais', N'MG')
INSERT [dbo].[Estados] ([ID_Estado], [Nome], [Sigla]) VALUES (14, N'Pará', N'PA')
INSERT [dbo].[Estados] ([ID_Estado], [Nome], [Sigla]) VALUES (15, N'Paraíba', N'PB')
INSERT [dbo].[Estados] ([ID_Estado], [Nome], [Sigla]) VALUES (16, N'Paraná', N'PR')
INSERT [dbo].[Estados] ([ID_Estado], [Nome], [Sigla]) VALUES (17, N'Pernambuco', N'PE')
INSERT [dbo].[Estados] ([ID_Estado], [Nome], [Sigla]) VALUES (18, N'Piauí', N'PI')
INSERT [dbo].[Estados] ([ID_Estado], [Nome], [Sigla]) VALUES (19, N'Rio de Janeiro', N'RJ')
INSERT [dbo].[Estados] ([ID_Estado], [Nome], [Sigla]) VALUES (20, N'Rio Grande do Norte', N'RN')
INSERT [dbo].[Estados] ([ID_Estado], [Nome], [Sigla]) VALUES (21, N'Rio Grande do Sul', N'RS')
INSERT [dbo].[Estados] ([ID_Estado], [Nome], [Sigla]) VALUES (22, N'Rondônia', N'RO')
INSERT [dbo].[Estados] ([ID_Estado], [Nome], [Sigla]) VALUES (23, N'Roraima', N'RR')
INSERT [dbo].[Estados] ([ID_Estado], [Nome], [Sigla]) VALUES (24, N'Santa Catarina', N'SC')
INSERT [dbo].[Estados] ([ID_Estado], [Nome], [Sigla]) VALUES (25, N'São Paulo', N'SP')
INSERT [dbo].[Estados] ([ID_Estado], [Nome], [Sigla]) VALUES (26, N'Sergipe', N'SE')
INSERT [dbo].[Estados] ([ID_Estado], [Nome], [Sigla]) VALUES (27, N'Tocantins', N'TO')
SET IDENTITY_INSERT [dbo].[Estados] OFF
GO
SET IDENTITY_INSERT [dbo].[Transacoes] ON 

INSERT [dbo].[Transacoes] ([ID_Transacao], [Numero_CPF], [ID_Corretor], [ID_Cidade], [Ativo], [Data_Transacao]) VALUES (1, N'123.456.789-99', 3, 2, 1, CAST(N'2024-09-18T00:00:00.000' AS DateTime))
INSERT [dbo].[Transacoes] ([ID_Transacao], [Numero_CPF], [ID_Corretor], [ID_Cidade], [Ativo], [Data_Transacao]) VALUES (2, N'123.456.789-68', 4, 23, 1, CAST(N'2024-09-18T00:00:00.000' AS DateTime))
INSERT [dbo].[Transacoes] ([ID_Transacao], [Numero_CPF], [ID_Corretor], [ID_Cidade], [Ativo], [Data_Transacao]) VALUES (3, N'123.456.789-99', 4, 1, 0, CAST(N'2024-09-18T00:00:00.000' AS DateTime))
SET IDENTITY_INSERT [dbo].[Transacoes] OFF
GO
SET ANSI_PADDING ON
GO
/****** Object:  Index [UQ__Clientes__03BFAB76D1274668]    Script Date: 11/10/2024 09:51:30 ******/
ALTER TABLE [dbo].[Clientes] ADD UNIQUE NONCLUSTERED 
(
	[Numero_CPF] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
GO
ALTER TABLE [dbo].[Cidades]  WITH CHECK ADD FOREIGN KEY([ID_Estado])
REFERENCES [dbo].[Estados] ([ID_Estado])
GO
ALTER TABLE [dbo].[Transacoes]  WITH CHECK ADD FOREIGN KEY([ID_Cidade])
REFERENCES [dbo].[Cidades] ([ID_Cidade])
GO
ALTER TABLE [dbo].[Transacoes]  WITH CHECK ADD FOREIGN KEY([ID_Corretor])
REFERENCES [dbo].[Corretores] ([ID_Corretor])
GO
ALTER TABLE [dbo].[Transacoes]  WITH CHECK ADD FOREIGN KEY([Numero_CPF])
REFERENCES [dbo].[Clientes] ([Numero_CPF])
GO
/****** Object:  StoredProcedure [dbo].[GetHttpRequest]    Script Date: 11/10/2024 09:51:30 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- Parte 2: Criar procedimento GetHttpRequest
CREATE   PROCEDURE [dbo].[GetHttpRequest]
    @url NVARCHAR(MAX),
    @result NVARCHAR(MAX) OUTPUT
AS
BEGIN
    DECLARE @obj INT
    DECLARE @ret INT
    DECLARE @status NVARCHAR(32)
    DECLARE @statusText NVARCHAR(32)

    EXEC @ret = sp_OACreate 'MSXML2.XMLHTTP', @obj OUT
    IF @ret <> 0 
    BEGIN
        RAISERROR('Não foi possível criar o objeto', 16, 1)
        RETURN
    END

    EXEC @ret = sp_OAMethod @obj, 'open', NULL, 'GET', @url, 'false'
    EXEC @ret = sp_OAMethod @obj, 'send'

    EXEC @ret = sp_OAGetProperty @obj, 'status', @status OUT
    EXEC @ret = sp_OAGetProperty @obj, 'statusText', @statusText OUT
    
    IF @status <> '200' 
    BEGIN
        RAISERROR('HTTP Request failed: %s %s', 16, 1, @status, @statusText)
        SET @result = NULL
        RETURN
    END

    EXEC @ret = sp_OAGetProperty @obj, 'responseText', @result OUT

    EXEC @ret = sp_OADestroy @obj
END
GO
/****** Object:  StoredProcedure [dbo].[sp_AtualizarCliente]    Script Date: 11/10/2024 09:51:30 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

CREATE PROCEDURE [dbo].[sp_AtualizarCliente]
    @ID_Cliente INT,
    @Nome_Cliente VARCHAR(100),
    @Numero_CPF VARCHAR(16)
AS
BEGIN
    UPDATE Clientes
    SET Nome_Cliente = @Nome_Cliente, Numero_CPF = @Numero_CPF
    WHERE ID_Cliente = @ID_Cliente;
END;

GO
/****** Object:  StoredProcedure [dbo].[sp_AtualizarCorretor]    Script Date: 11/10/2024 09:51:30 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO



CREATE PROCEDURE [dbo].[sp_AtualizarCorretor]
    @ID_Corretor INT,
    @Nome_Corretor VARCHAR(100)
AS
BEGIN
    UPDATE Corretores
    SET Nome_Corretor = @Nome_Corretor
    WHERE ID_Corretor = @ID_Corretor;
END;

GO
/****** Object:  StoredProcedure [dbo].[sp_AtualizarTransacao]    Script Date: 11/10/2024 09:51:30 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO


CREATE PROCEDURE [dbo].[sp_AtualizarTransacao]
    @IdTransacao INT,
    @Numero_CPF VARCHAR(16),
	@IdCorretor INT,
	@IdCidade INT,
	@Ativo INT,
    @Data_Transacao DATETIME
AS
BEGIN
    UPDATE Transacoes
    SET Numero_CPF = @Numero_CPF, ID_Corretor = @IdCorretor, ID_Cidade = @IdCidade, Ativo = @Ativo, Data_Transacao = @Data_Transacao
    WHERE ID_Transacao = @IdTransacao;
END;

GO
/****** Object:  StoredProcedure [dbo].[sp_ExcluirCliente]    Script Date: 11/10/2024 09:51:30 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

CREATE PROCEDURE [dbo].[sp_ExcluirCliente]
    @ID_Cliente INT
AS
BEGIN
    DELETE FROM Clientes WHERE ID_Cliente = @ID_Cliente;
END;

GO
/****** Object:  StoredProcedure [dbo].[sp_ExcluirCorretor]    Script Date: 11/10/2024 09:51:30 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

CREATE PROCEDURE [dbo].[sp_ExcluirCorretor]
    @ID_Corretor INT
AS
BEGIN
    DELETE FROM Corretores WHERE ID_Corretor = @ID_Corretor;
END;

GO
/****** Object:  StoredProcedure [dbo].[sp_ExcluirTransacao]    Script Date: 11/10/2024 09:51:30 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

CREATE PROCEDURE [dbo].[sp_ExcluirTransacao]
    @ID_Transacao INT
AS
BEGIN
    DELETE FROM Transacoes WHERE ID_Transacao = @ID_Transacao;
END;

GO
/****** Object:  StoredProcedure [dbo].[sp_InserirCliente]    Script Date: 11/10/2024 09:51:30 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

CREATE PROCEDURE [dbo].[sp_InserirCliente]
    @Nome_Cliente VARCHAR(100),
    @Numero_CPF VARCHAR(16)
AS
BEGIN
    INSERT INTO Clientes (Nome_Cliente, Numero_CPF)
    VALUES (@Nome_Cliente, @Numero_CPF);
END;

GO
/****** Object:  StoredProcedure [dbo].[sp_InserirCorretor]    Script Date: 11/10/2024 09:51:30 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO



CREATE PROCEDURE [dbo].[sp_InserirCorretor]
    @Nome_Corretor VARCHAR(100)
AS
BEGIN
    INSERT INTO Corretores (Nome_Corretor)
    VALUES (@Nome_Corretor);
END;

GO
/****** Object:  StoredProcedure [dbo].[sp_InserirTransacao]    Script Date: 11/10/2024 09:51:30 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO



CREATE PROCEDURE [dbo].[sp_InserirTransacao]
    @Numero_CPF VARCHAR(16),
	@IdCorretor INT,
	@IdCidade INT,
	@Ativo INT,
    @Data_Transacao DATETIME
AS
BEGIN
    INSERT INTO Transacoes (Numero_CPF, ID_Corretor, ID_Cidade, Ativo, Data_Transacao)
    VALUES (@Numero_CPF, @IdCorretor, @IdCidade, @Ativo, @Data_Transacao);
END;

GO
