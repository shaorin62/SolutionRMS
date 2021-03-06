if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[web_board]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[web_board]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[web_tail]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[web_tail]
GO

CREATE TABLE [dbo].[web_board] (
	[board_num] [int] IDENTITY (1, 1) NOT NULL ,
	[b_num] [int] NOT NULL ,
	[name] [varchar] (25) COLLATE Korean_Wansung_CI_AS NOT NULL ,
	[email] [varchar] (50) COLLATE Korean_Wansung_CI_AS NULL ,
	[homepage] [varchar] (50) COLLATE Korean_Wansung_CI_AS NULL ,
	[title] [varchar] (80) COLLATE Korean_Wansung_CI_AS NOT NULL ,
	[content] [text] COLLATE Korean_Wansung_CI_AS NOT NULL ,
	[pwd] [varchar] (20) COLLATE Korean_Wansung_CI_AS NOT NULL ,
	[writeday] [varchar] (50) COLLATE Korean_Wansung_CI_AS NULL ,
	[sessions] [varchar] (30) COLLATE Korean_Wansung_CI_AS NULL ,
	[r_num] [int] NULL ,
	[readnum] [int] NULL ,
	[comment_count] [int] NULL ,
	[u_ip] [varchar] (50) COLLATE Korean_Wansung_CI_AS NULL ,
	[tag] [varchar] (5) COLLATE Korean_Wansung_CI_AS NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[web_tail] (
	[tail_num] [varchar] (10) COLLATE Korean_Wansung_CI_AS NOT NULL ,
	[name] [varchar] (25) COLLATE Korean_Wansung_CI_AS NOT NULL ,
	[pwd] [varchar] (25) COLLATE Korean_Wansung_CI_AS NOT NULL ,
	[email] [varchar] (50) COLLATE Korean_Wansung_CI_AS NULL ,
	[content] [text] COLLATE Korean_Wansung_CI_AS NOT NULL ,
	[writeday] [varchar] (50) COLLATE Korean_Wansung_CI_AS NULL ,
	[u_ip] [varchar] (30) COLLATE Korean_Wansung_CI_AS NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

