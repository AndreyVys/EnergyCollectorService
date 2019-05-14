CREATE DATABASE [energy]
GO
USE [energy]
GO

CREATE LOGIN energy 
	WITH PASSWORD='energy', 
	DEFAULT_DATABASE=[energy],  
	CHECK_EXPIRATION=OFF, 
	CHECK_POLICY=ON
GO


CREATE USER [energy] FOR LOGIN [energy] WITH DEFAULT_SCHEMA=[dbo]
GO

ALTER AUTHORIZATION ON DATABASE::energy TO energy
GO

ALTER ROLE [db_owner] ADD MEMBER [energy]
GO


ALTER AUTHORIZATION ON DATABASE::energy TO energy
GO

USE [energy]
GO

CREATE TABLE [dbo].[AccuEnergy](
	[id] [int] IDENTITY(1,1) NOT NULL,
	[datetime] [datetime] NULL,
	[Port] [varchar](6) NULL,
	[Address] [varchar](3) NULL,
	[SerialNumber] [int] NULL,
	[SummaryActiveEnergy] [int] NULL,
	[SummaryReactiveEnergy] [int] NULL,
	[YearActiveEnergy] [int] NULL,
	[YearReactiveEnergy] [int] NULL,
	[MonthActiveEnergy] [int] NULL,
	[MonthReactiveEnergy] [int] NULL,
	[DayActiveEnergy] [int] NULL,
	[DayReactiveEnergy] [int] NULL,
	[VoltageRatio] [int] NULL,
	[CurrentRatio] [int] NULL,
	[CollectorType] [int] NULL,
 CONSTRAINT [PK_AccuEnergy] PRIMARY KEY CLUSTERED 
(
	[id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO


CREATE TABLE [dbo].[Collectors](
	[id] [int] IDENTITY(1,1) NOT NULL,
	[SerialNumber] [int] NULL,
	[ComportName] [nchar](10) NULL,
	[Address] [nchar](3) NULL,
	[CollectorGroup] [int] NULL,
	[Description] [varchar](1024) NULL,
	[CollectorType] [int] NULL,
	[BaudRate] [int] NULL,
 CONSTRAINT [PK_Collectors] PRIMARY KEY CLUSTERED 
(
	[id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[CollectorsType](
	[id] [int] IDENTITY(1,1) NOT NULL,
	[CollectorType] [int] NOT NULL,
	[Description] [varchar](50) NULL,
 CONSTRAINT [PK_CollectorsType] PRIMARY KEY CLUSTERED 
(
	[id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO


CREATE TABLE [dbo].[CollectorGroups](
	[id] [int] IDENTITY(1,1) NOT NULL,
	[CollectorGroup] [int] NULL,
	[Description] [varchar](1024) NULL,
 CONSTRAINT [PK_CollectorGroups] PRIMARY KEY CLUSTERED 
(
	[id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO


