Attribute VB_Name = "MiscFunctions"
Option Explicit

Public Sub WriteToLogWindow(strString As String, Optional TimeStamp As Boolean)

    Dim strTimeStamp As String
    Dim tmpText As String

    If TimeStamp = True Then strTimeStamp = "[" & Now & "] "
    
    tmpText = frmMain.txtSvrLog.Text
    If Len(tmpText) > 20000 Then tmpText = Right$(tmpText, 20000)
    
    frmMain.txtSvrLog.Text = tmpText & vbCrLf & strTimeStamp & strString
    frmMain.txtSvrLog.SelStart = Len(frmMain.txtSvrLog.Text)

End Sub

Public Function StripNulls(strString As Variant) As String

    If InStr(strString, vbNullChar) Then
        StripNulls = Left(strString, InStr(strString, vbNullChar) - 1)
    Else
        StripNulls = strString
    End If

End Function
