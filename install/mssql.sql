if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[DiyPage_ACT]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[DiyPage_ACT]
-
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ACT_LabelFolder]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ACT_LabelFolder]
-
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[AC_ACT]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[AC_ACT]
-
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ATT_ACT]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ATT_ACT]
-
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Admin_Act]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Admin_Act]
-
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Article_ACT]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Article_ACT]
-
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Book_ACT]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Book_ACT]
-
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Card_ACT]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Card_ACT]
-
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ClassLink_Act]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ClassLink_Act]
-
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Class_Act]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Class_Act]
-
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Comment_Act]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Comment_Act]
-
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Show_Act]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Show_Act]
-
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Config_Act]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Config_Act]
-
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Digg_ACT]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Digg_ACT]
-
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[DiyMenu_ACT]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[DiyMenu_ACT]
-
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[DownType_ACT]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[DownType_ACT]
-
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Edays_ACT]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Edays_ACT]
-
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Friend_ACT]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Friend_ACT]
-
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Group_Act]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Group_Act]
-
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Label_Act]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Label_Act]
-
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Link_Act]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Link_Act]
-
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Log_Act]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Log_Act]
-
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Message_Act]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Message_Act]
-
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ModeForm_ACT]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ModeForm_ACT]
-
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ModeUser_Act]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ModeUser_Act]
-
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Mode_Act]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Mode_Act]
-
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Money_Log_ACT]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Money_Log_ACT]
-
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Plus_ACT]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Plus_ACT]
-
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Point_Log_ACT]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Point_Log_ACT]
-
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Sitelink_ACT]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Sitelink_ACT]
-
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Table_ACT]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Table_ACT]
-
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Tags_ACT]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Tags_ACT]
-
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Upload_Act]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Upload_Act]
-
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Field_User_ACT]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Field_User_ACT]
-
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[User_Act]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[User_Act]
-
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Vote_act]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Vote_act]
-
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ads]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ads]
-
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[space_ACT]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[space_ACT]
-
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[templets_act]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[templets_act]
-
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Special_ACT]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Special_ACT]
-
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[SpecialPicUrl_ACT]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[SpecialPicUrl_ACT]
-
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Node_ACT]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Node_ACT]
-
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Mood_Plus_ACT]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Mood_Plus_ACT]
-
if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Mood_List_ACT]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Mood_List_ACT]
-
CREATE TABLE [dbo].[Special_ACT] (
	[ID] Int IDENTITY (1, 1) NOT NULL PRIMARY KEY,
	[pubdate] [datetime] NULL ,
	[title] [nvarchar] (250),
	[tempurl] [nvarchar] (250),
	[filename] [nvarchar] (250),
	[PicIndex] [nvarchar] (250),
	[Content] [ntext],
	[Hits] [int] Default 0 ,
	[writer] [nvarchar] (250)
) ON [PRIMARY]
-
CREATE TABLE [dbo].[SpecialPicUrl_ACT] (
	[ID] Int IDENTITY (1, 1) NOT NULL PRIMARY KEY,
	[sid] [int] Default 0 ,
	[title] [nvarchar] (250),
	[picurl] [nvarchar] (100)
) ON [PRIMARY]
-
CREATE TABLE [dbo].[Node_ACT] (
	[ID] Int IDENTITY (1, 1) NOT NULL PRIMARY KEY,
	[ModeID] [int] Default 0 ,
	[ContentLen] [int] Default 0 ,
	[DateForm] [int] Default 0 ,
	[ListNumber] [int] Default 0 ,
	[SID] [int] Default 0 ,
	[AddDate] [datetime] NULL ,
	[arcid] [nvarchar] (250),
	[isauto] [nvarchar] (250),
	[keywords] [nvarchar] (250),
	[classid] [nvarchar] (250),
	[DiyContent] [ntext],
	[TitleLen] [int] Default 0 ,
	[notename] [nvarchar] (100)
) ON [PRIMARY]
-
CREATE TABLE [dbo].[DiyPage_ACT] (
	[ID] Int IDENTITY (1, 1) NOT NULL PRIMARY KEY,
	[pageName] [nvarchar] (50),
	[DiyPath] [nvarchar] (50),
	[Content] [ntext],
	[tempurl] [nvarchar] (250),
	[FileName] [nvarchar] (100)
) ON [PRIMARY]
-
CREATE TABLE [dbo].[ACT_LabelFolder] (
	[ID] Int IDENTITY (1, 1) NOT NULL PRIMARY KEY,
	[Foldername] [nvarchar] (50)
) ON [PRIMARY]
-
CREATE TABLE [dbo].[AC_ACT] (
	[ID] Int IDENTITY (1, 1) NOT NULL PRIMARY KEY,
	[Field1] [nvarchar] (50),
	[Field2] [nvarchar] (50),
	[Types] [int] Default 0 
) ON [PRIMARY]
-
CREATE TABLE [dbo].[ATT_ACT] (
	[ID] Int IDENTITY (1, 1) NOT NULL PRIMARY KEY,
	[AID] [int] Default 0 ,
	[Aname] [nvarchar] (50)
) ON [PRIMARY]
-
CREATE TABLE [dbo].[Admin_Act] (
	[ID] Int IDENTITY (1, 1) NOT NULL PRIMARY KEY,
	[Admin_Name] [nvarchar] (50),
	[PassWord] [nvarchar] (50),
	[User_Name] [nvarchar] (50),
	[RealName] [nvarchar] (50),
	[Sex] [tinyint] Default 0 ,
	[TEL] [nvarchar] (50),
	[Email] [nvarchar] (50),
	[AddDate] [datetime] NULL ,
	[LoginTime] [datetime] NULL ,
	[LoginIP] [nvarchar] (50),
	[Locked] [tinyint] Default 0 ,
	[SuperTF] [tinyint] Default 0 ,
	[LoginNumber] [int] Default 0 ,
	[Description] [ntext],
	[Purview] [nvarchar] (200),
	[ACTCMS_QXLX] [ntext],
	[ACT_Other] [ntext]
) ON [PRIMARY]
-
CREATE TABLE [dbo].[Article_ACT] (
	[ID] Int IDENTITY (1, 1) NOT NULL PRIMARY KEY,
	[ClassID] [nvarchar] (20),
	[Title] [nvarchar] (100) ,
	[IntactTitle] [nvarchar] (250),
	[Intro] [ntext],
	[Content] [ntext],
	[Hits] [int] Default 0 ,
	[rev] [tinyint] Default 0 ,
	[ChargeType] [tinyint] Default 0 ,
	[InfoPurview] [tinyint] Default 0 ,
	[KeyWords] [nvarchar] (250),
	[author] [nvarchar] (250),
	[CopyFrom] [nvarchar] (250),
	[UpdateTime] [datetime] NULL ,
	[TemplateUrl] [nvarchar] (50),
	[FileName] [nvarchar] (200),
	[isAccept] [tinyint] Default 0 ,
	[delif] [tinyint] Default 0 ,
	[ArticleInput] [nvarchar] (250),
	[Slide] [tinyint] Default 0 ,
	[PicUrl] [nvarchar] (250),
	[Ismake] [tinyint] Default 0 ,
	[Digg] [int] Default 0 ,
	[down] [int] Default 0 ,
	[ReadPoint] [int] Default 0 ,
	[PitchTime] [int] Default 0 ,
	[ReadTimes] [int] Default 0 ,
	[DividePercent] [int] Default 0 ,
	[OrderID] [int] Default 0 ,
	[commentscount] [int] Default 0 ,
	[arrGroupID] [nvarchar] (250),
	[ATT] [smallint] Default 0 ,
	[IStop] [int] Default 0 ,
	[userid] [int] Default 0 ,
	[ActLink] [tinyint] Default 0
) ON [PRIMARY]
-
CREATE TABLE [dbo].[Book_ACT] (
	[ID] Int IDENTITY (1, 1) NOT NULL PRIMARY KEY,
	[show] [nvarchar] (50),
	[name] [nvarchar] (50),
	[qq] [nvarchar] (50),
	[mail] [nvarchar] (50),
	[url] [nvarchar] (50),
	[xq] [smallint] NULL ,
	[nr] [ntext],
	[addtime] [datetime] NULL ,
	[hf] [ntext],
	[ip] [nvarchar] (50),
	[sh] [tinyint] Default 0 
) ON [PRIMARY]
-
CREATE TABLE [dbo].[Card_ACT] (
	[ID] Int IDENTITY (1, 1) NOT NULL PRIMARY KEY,
	[CardNum] [nvarchar] (50),
	[CardPass] [nvarchar] (50),
	[title] [nvarchar] (250),
	[allgroupid] [nvarchar] (50),
	[Money] [money] NULL ,
	[ValidNum] [int] Default 0 ,
	[ValidUnit] [int] Default 0 ,
	[AddDate] [datetime] NULL ,
	[EndDate] [datetime] NULL ,
	[UseDate] [datetime] NULL ,
	[UserID] [int] Default 0 ,
	[IsUsed] [int] Default 0 ,
	[grGroupID] [int] Default 0 ,
	[ExpireGroupID] [int] Default 0 ,
	[IsSale] [int] Default 0 
) ON [PRIMARY]
-
CREATE TABLE [dbo].[ClassLink_Act] (
	[ID] Int IDENTITY (1, 1) NOT NULL PRIMARY KEY,
	[ClassLinkName] [nvarchar] (100),
	[Description] [ntext],
	[AddDate] [datetime] NULL 
) ON [PRIMARY]
-
CREATE TABLE [dbo].[Class_Act] (
	[ID] Int IDENTITY (1, 1) NOT NULL PRIMARY KEY,
	[ModeID] [smallint] NULL ,
	[ClassID] [nvarchar] (50),
	[OrderID] [smallint] NULL ,
	[enname] [nvarchar] (200),
	[ClassName] [nvarchar] (100),
	[ClassEName] [nvarchar] (100),
	[ParentID] [nvarchar] (15),
	[FolderTemplate] [nvarchar] (100),
	[ConTentTemplate] [nvarchar] (50),
	[Extension] [nvarchar] (50),
	[ClassKeywords] [ntext],
	[ClassDescription] [ntext],
	[ACTlink] [int] Default 0 ,
	[ClassPurview] [int] Default 0 ,
	[ClassReadPoint] [int] Default 0 ,
	[ClassChargeType] [int] Default 0 ,
	[ClassPitchTime] [int] Default 0 ,
	[ClassReadTimes] [int] Default 0 ,
	[ClassDividePercent] [int] Default 0 ,
	[ClassArrGroupID] [nvarchar] (250),
	[tg] [tinyint] Default 0 ,
	[dh] [tinyint] Default 0 ,
	[TGGroupID] [nvarchar] (255),
	[GroupIDClass] [nvarchar] (50),
	[moresite] [tinyint] Default 0 ,
	[labelfor] [tinyint] Default 0 ,
	[siteurl] [nvarchar] (250),
	[sitepath] [nvarchar] (250),
	[seotitle] [nvarchar] (250),
	[ClassPicUrl] [nvarchar] (250),
	[FilePathName] [nvarchar] (100),
	[content] [ntext],
	[makehtmlname] [nvarchar] (100),
	[pageTemplate] [nvarchar] (50)
) ON [PRIMARY]
-
CREATE TABLE [dbo].[Comment_Act] (
	[ID] Int IDENTITY (1, 1) NOT NULL PRIMARY KEY,
	[ModeID] [int] Default 0 ,
	[ClassID] [nvarchar] (50),
	[acticleID] [int] Default 0 ,
	[Email] [nvarchar] (50),
	[UserIP] [nvarchar] (50),
	[Content] [ntext],
	[Locked] [tinyint] Default 0 ,
	[AddDate] [datetime] NULL ,
	[userid] [int] Default 0 ,
	[Y] [int]  Default 0  ,
	[N] [int]  Default 0
) ON [PRIMARY]
-
CREATE TABLE [dbo].[Show_Act] (
	[ID] Int IDENTITY (1, 1) NOT NULL PRIMARY KEY,
	[ModeID] [int] Default 0 ,
	[acticleID] [int] Default 0 ,
	[AddDate] [datetime] Default 0 ,
	[userid] [int] Default 0
) ON [PRIMARY]
-
CREATE TABLE [dbo].[Config_Act] (
	[ID] Int IDENTITY (1, 1) NOT NULL PRIMARY KEY,
	[ActCMS_SysSetting] [ntext],
	[ActCMS_OtherSetting] [ntext],
	[ActCMS_Theme] varchar(250),
	[ActCMS_Upfile] [ntext]
) ON [PRIMARY]
-
CREATE TABLE [dbo].[Digg_ACT] (
	[ID] Int IDENTITY (1, 1) NOT NULL PRIMARY KEY,
	[IP] [nvarchar] (50),
	[NewsID] [int] Default 0 ,
	[DiggTime] [datetime] NULL ,
	[Digg] [tinyint] Default 0 ,
	[ModeID] [smallint] Default 0 ,
	[users] [nvarchar] (50)
) ON [PRIMARY]
-
CREATE TABLE [dbo].[DiyMenu_ACT] (
	[ID] Int IDENTITY (1, 1) NOT NULL PRIMARY KEY,
	[MenuName] [nvarchar] (50),
	[MenuUrl] [nvarchar] (200),
	[OpenWay] [nvarchar] (50),
	[AdminID] [smallint] Default 0 
) ON [PRIMARY]
-
CREATE TABLE [dbo].[DownType_ACT] (
	[ID] Int IDENTITY (1, 1) NOT NULL PRIMARY KEY,
	[rootid] [int] Default 0 ,
	[DownName] [nvarchar] (50),
	[IFLock] [tinyint] Default 0 ,
	[DownPath] [nvarchar] (255),
	[UserGroup] [nvarchar] (50),
	[DownPoint] [nvarchar] (50),
	[isDisp] [tinyint] Default 0 ,
	[IsOuter] [int] Default 0 
) ON [PRIMARY]


CREATE TABLE [dbo].[Edays_ACT] (
	[ID] Int IDENTITY (1, 1) NOT NULL PRIMARY KEY,
	[UserID] [int] Default 0 ,
	[Edays] [int] Default 0 ,
	[Flag] [int] Default 0 ,
	[AddDate] [datetime] NULL ,
	[IP] [nvarchar] (20),
	[UserLog] [nvarchar] (255),
	[Descript] [nvarchar] (255)
) ON [PRIMARY]


CREATE TABLE [dbo].[Friend_ACT] (
	[ID] Int IDENTITY (1, 1) NOT NULL PRIMARY KEY,
	[AddDate] [datetime] NULL ,
	[Userid] [int] Default 0 ,
	[flag] [tinyint] Default 0 ,
	[UM] [smallint] Default 0 ,
	[U] [int] Default 0 
) ON [PRIMARY]


CREATE TABLE [dbo].[Group_Act] (
	[GroupID] Int IDENTITY (1, 1) NOT NULL PRIMARY KEY,
	[DefaultGroup] [tinyint] Default 0 ,
	[Description] [ntext],
	[ChargeType] [tinyint] Default 0 ,
	[GroupPoint] [int] Default 0 ,
	[ValidDays] [int] Default 0 ,
	[ModeID] [smallint] Default 0 ,
	[GroupSetting] [ntext],
	[GroupName] [nvarchar] (50)
) ON [PRIMARY]


CREATE TABLE [dbo].[Label_Act] (
	[ID] Int IDENTITY (1, 1) NOT NULL PRIMARY KEY,
	[LabelName] [nvarchar] (50) ,
	[LabelContent] [ntext],
	[Description] [ntext],
	[LabelType] [tinyint] Default 0 ,
	[LabelFlag] [tinyint] Default 0 ,
	[AddDate] [datetime] NULL 
) ON [PRIMARY]


CREATE TABLE [dbo].[Link_Act] (
	[ID] Int IDENTITY (1, 1) NOT NULL PRIMARY KEY,
	[ClassLinkID] [int] Default 0 ,
	[SiteName] [nvarchar] (255),
	[Webadmin] [nvarchar] (50),
	[Email] [nvarchar] (50),
	[Url] [nvarchar] (255),
	[LinkType] [int] Default 0 ,
	[Logo] [nvarchar] (150),
	[Description] [ntext],
	[Rec] [int] Default 0 ,
	[AddDate] [datetime] NULL ,
	[Locked] [tinyint] Default 0 ,
	[sh] [int] Default 0 
) ON [PRIMARY]
-
CREATE TABLE [dbo].[Log_Act] (
	[ID] Int IDENTITY (1, 1) NOT NULL PRIMARY KEY,
	[UserName] [nvarchar] (50),
	[ACT] [int] Default 0 ,
	[LoginIP] [nvarchar] (50),
	[Times] [datetime] NULL ,
	[ACTError] ntext null ,
	[GetHttp]  ntext null ,
) ON [PRIMARY]


CREATE TABLE [dbo].[Message_Act] (
	[ID] Int IDENTITY (1, 1) NOT NULL PRIMARY KEY,
	[Title] [nvarchar] (100),
	[Content] [ntext],
	[Flag] [tinyint] Default 0 ,
	[Sendtime] [datetime] NULL ,
	[userid] [int] Default 0 ,
	[UM] [smallint] Default 0 ,
	[U] [int] Default 0 
) ON [PRIMARY]


CREATE TABLE [dbo].[ModeForm_ACT] (
	[ModeID] Int IDENTITY (1, 1) NOT NULL PRIMARY KEY,
	[ModeName] [nvarchar] (50),
	[ModeTable] [nvarchar] (50),
	[ModeNote] [nvarchar] (50),
	[UploadPath] [nvarchar] (50),
	[UploadSize] [nvarchar] (50),
	[UnlockTime] [tinyint] Default 0 ,
	[StartTime] [datetime] NULL ,
	[EndTime] [datetime] NULL ,
	[UserGroupList] [nvarchar] (50),
	[SubmitNum] [tinyint] Default 0 ,
	[Moneys] [smallint] NULL ,
	[FormCode] [tinyint] Default 0 ,
	[IsMail] [tinyint] Default 0 ,
	[ModeStatus] [tinyint] Default 0 
) ON [PRIMARY]
-
CREATE TABLE [dbo].[ModeUser_Act] (
	[ModeID] Int IDENTITY (1, 1) NOT NULL PRIMARY KEY,
	[ModeName] [nvarchar] (50),
	[ModeTable] [nvarchar] (50),
	[ModeNote] [nvarchar] (50),
	[Template] [nvarchar] (50),
	[RegCode] [tinyint] Default 0 ,
	[SpaceID] [int] Default 0 
) ON [PRIMARY]
-
CREATE TABLE [dbo].[Mode_Act] (
	[ModeID] Int IDENTITY (1, 1) NOT NULL PRIMARY KEY,
	[ModeName] [nvarchar] (50),
	[IFmake] [int] Default 0 ,
	[ModeTable] [nvarchar] (50),
	[FileFolder] [tinyint] Default 0 ,
	[AutoPage] [int] Default 0 ,
	[ProjectUnit] [nvarchar] (50),
	[UpFilesDir] [nvarchar] (100),
	[ContentExtension] [nvarchar] (6),
	[ModeStatus] [tinyint] Default 0 ,
	[RefreshFlag] [int] Default 0 ,
	[RecyleIF] [int] Default 0 ,
	[ACT_DiY] [ntext],
	[MakeFolderDir] [nvarchar] (100),
	[WriteComment] [tinyint] Default 0 ,
	[CommentCode] [tinyint] Default 0 ,
	[Commentsize] [smallint] NULL ,
	[ModeNote] [nvarchar] (250),
	[CommentTemp] [nvarchar] (100),
	[adminmb] [tinyint] Default 0 ,
	[usermb] [tinyint] Default 0 
) ON [PRIMARY]
-
CREATE TABLE [dbo].[Money_Log_ACT] (
	[ID] Int IDENTITY (1, 1) NOT NULL PRIMARY KEY,
	[UserName] [nvarchar] (50),
	[ClientName] [nvarchar] (250),
	[Money] [money] NULL ,
	[CurrMoney] [money] NULL ,
	[MoneyType] [int] Default 0 ,
	[IncomeFlag] [int] Default 0 ,
	[ModeID] [int] Default 0 ,
	[InfoID] [int] Default 0 ,
	[OrderID] [nvarchar] (30),
	[PaymentID] [int] Default 0 ,
	[Remark] [nvarchar] (255),
	[PayTime] [datetime] NULL ,
	[LogTime] [datetime] NULL ,
	[Inputer] [nvarchar] (50),
	[IP] [nvarchar] (50)
) ON [PRIMARY]
-
CREATE TABLE [dbo].[Plus_ACT] (
	[ID] Int IDENTITY (1, 1) NOT NULL PRIMARY KEY,
	[PlusName] [nvarchar] (50),
	[PlusID] [nvarchar] (50),
	[PlusUrl] [nvarchar] (50),
	[IsUse] [int] Default 0 ,
	[PlusIntro] [nvarchar] (200),
	[OrderID] [smallint] NULL ,
	[PlusConfig] [ntext]
) ON [PRIMARY]
-
CREATE TABLE [dbo].[Point_Log_ACT] (
	[ID] Int IDENTITY (1, 1) NOT NULL PRIMARY KEY,
	[UserID] [nvarchar] (50),
	[ModeID] [int] Default 0 ,
	[InfoID] [int] Default 0 ,
	[ContributeFlag] [int] Default 0 ,
	[AddDate] [datetime] NULL ,
	[IP] [nvarchar] (50),
	[PointFlag] [int] Default 0 ,
	[CurrPoint] [int] Default 0 ,
	[Point] [int] Default 0 ,
	[Times] [int] Default 0 ,
	[UserLog] [nvarchar] (50),
	[Descript] [nvarchar] (255)
) ON [PRIMARY]


CREATE TABLE [dbo].[Sitelink_ACT] (
	[ID] Int IDENTITY (1, 1) NOT NULL PRIMARY KEY,
	[Title] [nvarchar] (50),
	[Url] [nvarchar] (50),
	[OpenType] [nvarchar] (50),
	[OrderID] [int] Default 0 ,
	[Num] [int] Default 0 ,
	[repset] [int] Default 0 ,
	[description] [ntext],
	[repcontent] [ntext],
	[IFS] [tinyint] Default 0 
) ON [PRIMARY]
-
CREATE TABLE [dbo].[Table_ACT] (
	[ID] Int IDENTITY (1, 1) NOT NULL PRIMARY KEY,
	[ModeID] [smallint] NULL ,
	[FieldName] [nvarchar] (50),
	[Title] [nvarchar] (50),
	[IsNotNull] [tinyint] Default 0 ,
	[OrderID] [smallint] NULL ,
	[Description] [nvarchar] (250),
	[FieldType] [nvarchar] (50),
	[Type_Default] [ntext],
	[width] [smallint] NULL ,
	[height] [smallint] NULL ,
	[Content] [ntext],
	[Type_Type] [int] Default 0 ,
	[ISType] [int] Default 0 ,
	[regEx] [nvarchar] (255),
	[regError] [nvarchar] (255),
	[SearchIF] [tinyint] Default 0 ,
	[ValueOnly] [tinyint] Default 0 ,
	[check] [tinyint] Default 0 ,
	[ACTCMS] [int] Default 0 
) ON [PRIMARY]
-
CREATE TABLE [dbo].[Tags_ACT] (
	[ID] Int IDENTITY (1, 1) NOT NULL PRIMARY KEY,
	[TagsChar] [nvarchar] (50),
	[ModeID] [tinyint] Default 0 ,
	[AddTime] [datetime] NULL ,
	[Hits] [int] Default 0 ,
	[ClicksTime] [datetime] NULL 
) ON [PRIMARY]
-
CREATE TABLE [dbo].[Upload_Act] (
	[ID] Int IDENTITY (1, 1) NOT NULL PRIMARY KEY,
	[ArtileID] [int] Default 0 ,
	[UpfileDir] [nvarchar] (250),
	[Extension] [nvarchar] (250),
	[UpdateTime] [datetime] NULL ,
	[ModeID] [smallint] NULL 
) ON [PRIMARY]
-
CREATE TABLE [dbo].[Field_User_ACT] (
	[ID] Int IDENTITY (1, 1) NOT NULL PRIMARY KEY,
	[UserID] [int] Default 0 
) ON [PRIMARY]
-
CREATE TABLE [dbo].[User_Act] (
	[UserID] Int IDENTITY (1, 1) NOT NULL PRIMARY KEY,
	[GroupID] [smallint] NULL ,
	[UserName] [nvarchar] (50),
	[PassWord] [nvarchar] (20),
	[LoginNumber] [int] Default 0 ,
	[ChargeType] [int] Default 0 ,
	[Score] [int] Default 0 ,
	[UModeID] [int] Default 0 ,
	[EDays] [int] Default 0 ,
	[Point] [int] Default 0 ,
	[Locked] [tinyint] Default 0 ,
	[Loginip] [nvarchar] (50),
	[RegDate] [datetime] NULL ,
	[BeginDate] [datetime] NULL ,
	[LoginTime] [datetime] NULL ,
	[Email] [nvarchar] (50),
	[Money] [money]  Default 0 ,
	[RealName] [nvarchar] (50),
	[sex] [tinyint] Default 0 ,
	[Province] [nvarchar] (100),
	[City] [nvarchar] (50),
	[Birthday] [nvarchar] (50),
	[Privacy] [int] Default 0 ,
	[CheckNum] [nvarchar] (16),
	[note] [nvarchar] (255),
	[ArticleNum] [int] Default 0 ,
	[QQ] [nvarchar] (15),
	[MSN] [nvarchar] (50),
	[address] [nvarchar] (255),
	[HomeTel] [nvarchar] (50),
	[Mobile] [nvarchar] (20),
	[postcode] [nvarchar] (20),
	[myface] [nvarchar] (100),
	[Question] [nvarchar] (250),
	[Answer] [nvarchar] (250),
	[templetsid] [smallint] NULL 
) ON [PRIMARY]
-
CREATE TABLE [dbo].[Vote_act] (
	[ID] Int IDENTITY (1, 1) NOT NULL PRIMARY KEY,
	[title] [nvarchar] (150),
	[isLock] [smallint] NULL ,
	[VoteTime] [datetime] NULL ,
	[VoteType] [tinyint] Default 0 ,
	[VoteNum] [int] Default 0 ,
	[rootid] [smallint] NULL ,
	[VoteStart] [datetime] NULL ,
	[VoteEnd] [datetime] NULL 
) ON [PRIMARY]
-
CREATE TABLE [dbo].[ads] (
	[ID] Int IDENTITY (1, 1) NOT NULL PRIMARY KEY,
	[ADID] [nvarchar] (20),
	[ADType] [smallint] NULL ,
	[ADSrc] [ntext],
	[ADCode] [ntext],
	[ADWidth] [smallint] NULL ,
	[ADHeight] [smallint] NULL ,
	[ADLink] [nvarchar] (150),
	[ADAlt] [nvarchar] (100),
	[ADNote] [nvarchar] (100),
	[ADViews] [int] Default 0 ,
	[ADHits] [int] Default 0 ,
	[ADStopViews] [int] Default 0 ,
	[ADStopHits] [int] Default 0 ,
	[ADStopDate] [datetime] NULL 
) ON [PRIMARY]-
CREATE TABLE [dbo].[Mood_Plus_ACT] (
	[ID] Int IDENTITY (1, 1) NOT NULL PRIMARY KEY,
	[Title] [nvarchar] (250),
	[Status] [smallint] Default 0 ,
	[TitleContent] [ntext],
	[PicContent] [ntext],
	[StartTime] [datetime] NULL ,
	[EndTime] [datetime] NULL ,
	[SubmitNum] [smallint] Default 0 ,
	[UnlockTime] [smallint] Default 0 
) ON [PRIMARY]-
CREATE TABLE [dbo].[Mood_List_ACT] (
	[ID] Int IDENTITY (1, 1) NOT NULL PRIMARY KEY,
	[Title] [nvarchar] (250),
	[ModeID] [int] Default 0 ,
	[MDID] [int] Default 0 ,
	[AID] [int] Default 0 ,
	[CID] [int] Default 0 ,
	[M0] [int] Default 0 ,
	[M1] [int] Default 0 ,
	[M2] [int] Default 0 ,
	[M3] [int] Default 0 ,
	[M4] [int] Default 0 ,
	[M5] [int] Default 0 ,
	[M6] [int] Default 0 ,
	[M7] [int] Default 0 ,
	[M8] [int] Default 0 ,
	[M9] [int] Default 0 ,
	[M10] [int] Default 0 ,
	[M11] [int] Default 0 ,
	[M12] [int] Default 0 ,
	[M13] [int] Default 0 ,
	[M14] [int] Default 0 ,
	[Hits] [int] Default 0
) ON [PRIMARY]
-
CREATE TABLE [dbo].[space_ACT] (
	[ID] Int IDENTITY (1, 1) NOT NULL PRIMARY KEY,
	[ClassName] [nvarchar] (50),
	[ClassOrder] [int] Default 0 ,
	[ClassTemp] [nvarchar] (50),
	[ModeID] [smallint] NULL ,
	[UModeID] [int] Default 0 
) ON [PRIMARY]
-
CREATE TABLE [dbo].[templets_act] (
	[ID] Int IDENTITY (1, 1) NOT NULL PRIMARY KEY,
	[templets] [nvarchar] (50),
	[UserSet] [nvarchar] (50)
) ON [PRIMARY] 