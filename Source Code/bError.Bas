Attribute VB_Name = "bError"
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

' Generic Error routine
Public Sub afx_GenericError(sProcedure As String, lErr As Long)
    On Error GoTo Err_afx_GenericError
    
    Beep
    MsgBox sProcedure & " " & Error$(lErr) & ". " & CStr(lErr)

Exit_afx_GenericError:
    Exit Sub
Err_afx_GenericError:
    Call afx_GenericError("afx_GenericError:", Err)
    Stop
End Sub
