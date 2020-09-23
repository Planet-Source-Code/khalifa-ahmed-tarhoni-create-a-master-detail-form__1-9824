/* Microsoft SQL Server - Scripting			*/
/* Server: SSHAFIQ					*/
/* Database: dbHotel					*/
/* Creation Date 7/15/00 5:12:08 PM 			*/

/****** Object:  Trigger dbo.trg_Update_Rooms    Script Date: 7/15/00 5:12:08 PM ******/
if exists (select * from sysobjects where id = object_id('dbo.trg_Update_Rooms') and sysstat & 0xf = 8)
	drop trigger dbo.trg_Update_Rooms
GO

/****** Object:  Trigger dbo.Web_784005824_1    Script Date: 7/15/00 5:12:08 PM ******/
if exists (select * from sysobjects where id = object_id('dbo.Web_784005824_1') and sysstat & 0xf = 8)
	drop trigger dbo.Web_784005824_1
GO

/****** Object:  Trigger dbo.Web_784005824_2    Script Date: 7/15/00 5:12:08 PM ******/
if exists (select * from sysobjects where id = object_id('dbo.Web_784005824_2') and sysstat & 0xf = 8)
	drop trigger dbo.Web_784005824_2
GO

/****** Object:  Trigger dbo.Web_784005824_3    Script Date: 7/15/00 5:12:08 PM ******/
if exists (select * from sysobjects where id = object_id('dbo.Web_784005824_3') and sysstat & 0xf = 8)
	drop trigger dbo.Web_784005824_3
GO

/****** Object:  Stored Procedure dbo.proc_RoomDesc    Script Date: 7/15/00 5:12:08 PM ******/
if exists (select * from sysobjects where id = object_id('dbo.proc_RoomDesc') and sysstat & 0xf = 4)
	drop procedure dbo.proc_RoomDesc
GO

/****** Object:  View dbo.vr_Reservations    Script Date: 7/15/00 5:12:08 PM ******/
if exists (select * from sysobjects where id = object_id('dbo.vr_Reservations') and sysstat & 0xf = 2)
	drop view dbo.vr_Reservations
GO

/****** Object:  Table dbo.tbl_Hotels    Script Date: 7/15/00 5:12:08 PM ******/
if exists (select * from sysobjects where id = object_id('dbo.tbl_Hotels') and sysstat & 0xf = 3)
	drop table dbo.tbl_Hotels
GO

/****** Object:  Table dbo.tbl_PaymentTypes    Script Date: 7/15/00 5:12:08 PM ******/
if exists (select * from sysobjects where id = object_id('dbo.tbl_PaymentTypes') and sysstat & 0xf = 3)
	drop table dbo.tbl_PaymentTypes
GO

/****** Object:  Table dbo.tbl_Reservations    Script Date: 7/15/00 5:12:08 PM ******/
if exists (select * from sysobjects where id = object_id('dbo.tbl_Reservations') and sysstat & 0xf = 3)
	drop table dbo.tbl_Reservations
GO

/****** Object:  Table dbo.tbl_Rooms    Script Date: 7/15/00 5:12:08 PM ******/
if exists (select * from sysobjects where id = object_id('dbo.tbl_Rooms') and sysstat & 0xf = 3)
	drop table dbo.tbl_Rooms
GO

/****** Object:  Table dbo.tbl_RoomTypes    Script Date: 7/15/00 5:12:08 PM ******/
if exists (select * from sysobjects where id = object_id('dbo.tbl_RoomTypes') and sysstat & 0xf = 3)
	drop table dbo.tbl_RoomTypes
GO

/****** Object:  Table dbo.tbl_Hotels    Script Date: 7/15/00 5:12:08 PM ******/
CREATE TABLE dbo.tbl_Hotels (
	Hotel int NOT NULL ,
	Name varchar (50) NULL 
)
GO

/****** Object:  Table dbo.tbl_PaymentTypes    Script Date: 7/15/00 5:12:08 PM ******/
CREATE TABLE dbo.tbl_PaymentTypes (
	Payment int NOT NULL ,
	Description varchar (50) NULL 
)
GO

/****** Object:  Table dbo.tbl_Reservations    Script Date: 7/15/00 5:12:08 PM ******/
CREATE TABLE dbo.tbl_Reservations (
	ResNo int NOT NULL ,
	LastName varchar (25) NULL ,
	FirstName varchar (20) NULL ,
	Address varchar (50) NULL ,
	City varchar (30) NULL ,
	State varchar (2) NULL ,
	Postal varchar (10) NULL ,
	Phone varchar (15) NULL ,
	Payment int NULL ,
	Amount money NULL ,
	Hotel int NOT NULL ,
	Room smallint NULL ,
	DateIn datetime NULL ,
	DateOut datetime NULL ,
	DateNow datetime NULL 
)
GO

/****** Object:  Table dbo.tbl_Rooms    Script Date: 7/15/00 5:12:08 PM ******/
CREATE TABLE dbo.tbl_Rooms (
	Hotel int NOT NULL ,
	Room smallint NOT NULL ,
	RoomType int NULL ,
	Price money NULL ,
	Comments text NULL ,
	RoomStatus varchar (1) NULL 
)
GO

/****** Object:  Table dbo.tbl_RoomTypes    Script Date: 7/15/00 5:12:08 PM ******/
CREATE TABLE dbo.tbl_RoomTypes (
	RoomType int NOT NULL ,
	Beds smallint NULL ,
	Description varchar (254) NULL 
)
GO

/****** Object:  View dbo.vr_Reservations    Script Date: 7/15/00 5:12:08 PM ******/
/****** Object:  View dbo.vr_Reservations    Script Date: 1/20/97 2:28:56 AM ******/
CREATE VIEW vr_Reservations
(ResNo, LastName, FirstName, PaymentDesc, Room, Beds, RoomDesc)
AS
SELECT A.ResNo, A.LastName, A.FirstName
     , B.Description
     , C.Room
     , D.Beds, D.Description
  FROM tbl_Reservations A, tbl_PaymentTypes B
     , tbl_Rooms C, tbl_RoomTypes D
 WHERE B.Payment = A.Payment
   AND C.Room = A.Room
   AND C.Hotel = A.Hotel
   AND D.RoomType = C.RoomType

GO

/****** Object:  Stored Procedure dbo.proc_RoomDesc    Script Date: 7/15/00 5:12:08 PM ******/
/****** Object:  Stored Procedure dbo.proc_RoomDesc    Script Date: 1/20/97 2:28:56 AM ******/
CREATE PROCEDURE proc_RoomDesc @Hotel INT AS
SELECT a.room, b.beds, b.description, a.price
  FROM tbl_Rooms a, tbl_RoomTypes b
 WHERE a.hotel = @Hotel
   AND a.roomtype = b.roomtype


GO

/****** Object:  Trigger dbo.trg_Update_Rooms    Script Date: 7/15/00 5:12:08 PM ******/
/****** Object:  Trigger dbo.trg_Update_Rooms    Script Date: 1/20/97 2:28:56 AM ******/
CREATE TRIGGER trg_Update_Rooms ON dbo.tbl_Rooms 
FOR UPDATE
AS
IF UPDATE(room)
BEGIN
  PRINT "Trigger trg_Update_Rooms was executed."
  UPDATE tbl_Rooms
  SET tbl_Rooms.room = inserted.room
  FROM inserted, deleted
  WHERE tbl_Rooms.room = deleted.room
    AND tbl_Rooms.hotel = deleted.hotel
END


GO

/****** Object:  Trigger dbo.Web_784005824_1    Script Date: 7/15/00 5:12:08 PM ******/
/****** Object:  Trigger dbo.Web_784005824_1    Script Date: 1/20/97 2:28:56 AM ******/
CREATE TRIGGER Web_784005824_1 ON dbo.tbl_Reservations FOR INSERT AS 
begin 
   exec sp_makewebtask @outputfile='d:\InetPub\wwwroot\sandbox\resv_rpt.htm',@query='select a.name as "Hotel", count(resno) as "Number of Reservations" from tbl_hotels a, tbl_reservations b where a.hotel = b.hotel and b.datein >= "01/01/97" and b.datein <= "12/31/97" group by a.name ',@fixedfont=1,@bold=0,@italic=0,@colheaders=1,@lastupdated=1,@HTMLheader=2,@username='dbo',@dbname='db_HotelSystem',@webpagetitle='Reservation Count',@resultstitle='Reservation Count for Current Year',@maketask=0,@rowcnt=0 
end

GO

/****** Object:  Trigger dbo.Web_784005824_2    Script Date: 7/15/00 5:12:08 PM ******/
/****** Object:  Trigger dbo.Web_784005824_2    Script Date: 1/20/97 2:28:56 AM ******/
CREATE TRIGGER Web_784005824_2 ON dbo.tbl_Reservations FOR UPDATE AS 
begin 
   exec sp_makewebtask @outputfile='e:\InetPub\wwwroot\sandbox\resv_rpt.htm',@query='select a.name as "Hotel", count(resno) as "Number of Reservations" from tbl_hotels a, tbl_reservations b where a.hotel = b.hotel and b.datein >= "01/01/97" and b.datein <= "12/31/97" group by a.name ',@fixedfont=1,@bold=0,@italic=0,@colheaders=1,@lastupdated=1,@HTMLheader=2,@username='dbo',@dbname='db_HotelSystem',@webpagetitle='Reservation Count',@resultstitle='Reservation Count for Current Year',@maketask=0,@rowcnt=0 
end

GO

/****** Object:  Trigger dbo.Web_784005824_3    Script Date: 7/15/00 5:12:08 PM ******/
/****** Object:  Trigger dbo.Web_784005824_3    Script Date: 1/20/97 2:28:56 AM ******/
CREATE TRIGGER Web_784005824_3 ON dbo.tbl_Reservations FOR DELETE AS 
begin 
   exec sp_makewebtask @outputfile='e:\InetPub\wwwroot\sandbox\resv_rpt.htm',@query='select a.name as "Hotel", count(resno) as "Number of Reservations" from tbl_hotels a, tbl_reservations b where a.hotel = b.hotel and b.datein >= "01/01/97" and b.datein <= "12/31/97" group by a.name ',@fixedfont=1,@bold=0,@italic=0,@colheaders=1,@lastupdated=1,@HTMLheader=2,@username='dbo',@dbname='db_HotelSystem',@webpagetitle='Reservation Count',@resultstitle='Reservation Count for Current Year',@maketask=0,@rowcnt=0 
end

GO

