VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmReservations 
   Caption         =   "Reservations"
   ClientHeight    =   4950
   ClientLeft      =   1380
   ClientTop       =   1740
   ClientWidth     =   6675
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4950
   ScaleWidth      =   6675
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   495
      Left            =   5040
      TabIndex        =   4
      Top             =   420
      Width           =   1215
   End
   Begin VB.Frame fraCurrentReservations 
      Caption         =   "Current Reservations"
      Height          =   1995
      Left            =   60
      TabIndex        =   2
      Top             =   2820
      Width           =   6495
      Begin MSFlexGridLib.MSFlexGrid grdResv 
         Height          =   1515
         Left            =   180
         TabIndex        =   6
         Top             =   300
         Width           =   6120
         _ExtentX        =   10795
         _ExtentY        =   2672
         AllowUserResizing=   1
      End
   End
   Begin VB.ComboBox cboHotel 
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   540
      Width           =   2115
   End
   Begin MSFlexGridLib.MSFlexGrid grdRoomInfo 
      Height          =   1515
      Left            =   240
      TabIndex        =   5
      Top             =   1260
      Width           =   6060
      _ExtentX        =   10689
      _ExtentY        =   2672
   End
   Begin VB.Label lblHotelInfo 
      Caption         =   "Hotel Information"
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   1020
      Width           =   1215
   End
   Begin VB.Label lblHotel 
      Caption         =   "Hotel:"
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   300
      Width           =   1215
   End
End
Attribute VB_Name = "frmReservations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Assign the Resv member to the Data Object
Private m_Resv As CDataObjects

Private Sub cboHotel_Click()
    
    Dim sSQL As String
    
    On Error GoTo Err_cboHotel_Click
    
    Screen.MousePointer = vbHourglass
    
    sSQL = "Select Room, Beds, Description "
    sSQL = sSQL & "From tbl_Rooms a, tbl_Roomtypes b "
    sSQL = sSQL & "Where a.roomtype = b.roomtype "
    sSQL = sSQL & "And Hotel = " & cboHotel.ListIndex + 1
    m_Resv.FillGrid grdRoomInfo, sSQL, "Room", "Beds", "Description"
    
    ' Set the SQL Statment for Reservations
    sSQL = "Select ResNo as 'Res No', LastName as 'Last Name', "
    sSQL = sSQL & "FirstName as 'First Name', Phone "
    sSQL = sSQL & "From tbl_Reservations "
    sSQL = sSQL & "Where hotel = " & cboHotel.ListIndex + 1
    m_Resv.FillGrid grdResv, sSQL, "Res. No", "Last Name", "First Name", "Phone"
    
    Screen.MousePointer = vbDefault
    
Exit_cboHotel_Click:
    Exit Sub
Err_cboHotel_Click:
    Call afx_GenericError("cboHotel_Click:", Err)
    Resume Exit_cboHotel_Click
End Sub

Private Sub cmdOK_Click()
    
    On Error GoTo err_cmdOK_Click
    
    Unload Me
    
Exit_cmdOK_Click:
    Exit Sub
err_cmdOK_Click:
    Call afx_GenericError("cmdOK_Click:", Err)
    Resume Exit_cmdOK_Click
End Sub

Private Sub Form_Load()
    
    Dim sSQL As String
    
    On Error GoTo Err_Form_Load
    
    Set m_Resv = New CDataObjects
    
    ' Set the SQL Statment for Hotels
    sSQL = "Select Hotel, Name from tbl_Hotels"
    
    m_Resv.FillControl cboHotel, sSQL, "hotel", "name"

        
Exit_Form_Load:
    Exit Sub
Err_Form_Load:
    Call afx_GenericError("Form_Load:", Err)
    Resume Exit_Form_Load
End Sub

Private Sub Form_Unload(Cancel As Integer)

    On Error GoTo Err_Form_Unload
    
    ' Clear the reference
    Set m_Resv = Nothing
    
Exit_Form_Unload:
    Exit Sub
Err_Form_Unload:
    Call afx_GenericError("Form_Unload:", Err)
    Resume Exit_Form_Unload
End Sub
