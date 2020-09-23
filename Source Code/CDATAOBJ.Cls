VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDataObjects"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'******************************************************
'Wali Systems Inc.
'
'Copyright (c) 2000 - All rights reserved.
'
'The author hereby allows anyone to copy and paste this
'code into his/her own application without any prior
'permission. Remember, this code cannot be resubmitted
'to anyother site.
'
'Contact:
'========
'
'For any help regarding this code, contact us
'at the following address:
'
'Email: ytarhoni@hotmail.com
'
'******************************************************

Option Explicit

' Database
Private m_env As rdoEnvironment

' Connection
Private m_con As rdoConnection


' Fill any list type control with the data
'
' Parameters:
'   cntl        list type control to fill
'   sSQL        SQL statement
'   sIDColumn   name of the column containing the record ID
'   sColumns    name of the columns of data to display
Public Sub FillControl(cntl As Control, sSQL As String, sIDColumn As String, ParamArray sColumns() As Variant)

    Dim sData As String
    Dim i As Integer
    Dim rs As rdoResultset

    On Error GoTo Err_FillControl

    ' Clear the list
    cntl.Clear

    ' open the resultset
    Set rs = m_con.OpenResultset(sSQL)
    
    ' Add each record until the end of the file
    ' until the end of file flag
    If (Not rs.BOF) Then
        Do Until rs.EOF
            ' Concatenate each desired column
            sData = ""
            For i = 0 To UBound(sColumns)
                sData = sData & " " & rs(CStr(sColumns(i)))
            Next
            ' Add the item to the list
            cntl.AddItem sData
            cntl.ItemData(cntl.NewIndex) = rs(sIDColumn)
    
            ' Move to the next row
            rs.MoveNext
        Loop
    End If
    ' Set control to highlight the first record
    ' in the list
    cntl.ListIndex = 0
    
Exit_FillControl:
    Exit Sub
Err_FillControl:
    Call afx_GenericError("FillControl:", Err)
    Resume Exit_FillControl
End Sub

Private Sub Class_Initialize()
    
    Dim sSQL As String
    
    On Error GoTo Err_Class_Initialize
    
    ' use the default environment
    Set m_env = rdoEnvironments(0)
    ' Open the connection
    Set m_con = m_env.OpenConnection(dsName:="dsnmdf", _
        Connect:="uid=sa;pwd=") 'Don't forget to make changes in the connect string

Exit_Class_Initialize:
    Exit Sub
    
Err_Class_Initialize:
    Call afx_GenericError("Class_Initialize:", Err)
    Resume Exit_Class_Initialize
End Sub

Private Sub Class_Terminate()
    
    On Error GoTo Err_Class_Terminate
    
    ' Close the resultset and database
    If Not (m_env Is Nothing) Then
        m_env.Close
    End If
    
    ' Clear the references
    Set m_env = Nothing

Exit_Class_Terminate:
    Exit Sub
Err_Class_Terminate:
    Call afx_GenericError("Class_Terminate:", Err)
    Resume Exit_Class_Terminate
End Sub


' Fill an unbound grid control with the data
'
' Parameters:
'   cntl        list type control to fill
'   sSQL        SQL statement
'   sColumns    name of the columns of data to display
Public Sub FillGrid(cntl As Control, sSQL As String, ParamArray sColumns() As Variant)

    Dim sData As String
    Dim i As Integer
    Dim j As Long
    Dim lMaxRows As Long
    Dim rs As rdoResultset
    Dim sFieldName As String
    
    On Error GoTo Err_FillGrid
    
    ' Open the resultset
    Set rs = m_con.OpenResultset(sSQL)
    
    ' Set up the of the grid
    With cntl
        .Cols = UBound(sColumns) + 1
        .Rows = cntl.Rows
    End With
    For i = 0 To cntl.Cols - 1
        cntl.ColWidth(i) = 1440
    Next i
    
    ' Fill in the Column headers from the
    ' Database column names
    If (Not rs.BOF) Then
        cntl.Row = 0
        For i = 0 To cntl.Cols - 1
            cntl.Col = i
            sFieldName = rs.rdoColumns(i).Name
            cntl.Text = sFieldName
        Next
        j = 1
        
        Do Until rs.EOF
            cntl.Row = j
            For i = 0 To UBound(sColumns)
                With cntl
                    .Col = i
                    .Text = rs(i)
                End With
            Next
            ' Get next row
            rs.MoveNext
            j = j + 1
            cntl.Rows = j + 1
        Loop
    End If
Exit_FillGrid:
    Exit Sub
Err_FillGrid:
    Call afx_GenericError("FillGrid:", Err)
    Resume Exit_FillGrid
End Sub


