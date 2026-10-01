USE [YOURDATABASE]
GO

/****** Object:  Table [dbo].[tblTipAngles]    Script Date: 10/1/2026 6:15:45 AM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[tblTipAngles](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[DepartmentID] [int] NOT NULL,
	[ProgramName] [varchar](2000) NOT NULL,
	[ProbeName] [varchar](255) NOT NULL,
	[TipName] [varchar](255) NOT NULL,
	[IsStillThere] [smallint] NOT NULL
) ON [PRIMARY]
GO

ALTER TABLE [dbo].[tblTipAngles] ADD  CONSTRAINT [DF_tblTipAngles_IsStillThere]  DEFAULT ((0)) FOR [IsStillThere]
GO


CREATE TABLE [dbo].[tblTipAngle_ImportRun](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[DepartmentID] [int] NOT NULL,
	[StartTime] [datetime] NOT NULL,
	[EndTime] [datetime] NULL
) ON [PRIMARY]
GO