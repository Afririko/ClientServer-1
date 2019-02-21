Attribute VB_Name = "koneksi"
Option Explicit
'----------------------------------------------------
Public cn As New ADODB.Connection
Public rs As New ADODB.Recordset
Public StrKon As String
Public StrSQL As String
Public server, port, database, user, password As String
'----------------------------------------------------

Public Sub connect()
On Error GoTo buat_koneksi_Error:
If cn.State = adStateOpen Then
    cn.Close
End If
server = GetSetting(App.EXEName, "x", "server")
port = GetSetting(App.EXEName, "x", "port")
database = GetSetting(App.EXEName, "x", "database")
user = GetSetting(App.EXEName, "x", "user")
password = GetSetting(App.EXEName, "x", "password")

StrKon = "DRIVER={MySQL ODBC 3.51 Driver};" _
        & "SERVER=" & server & ";" _
        & "DATABASE=" & database & ";" _
        & "port=" & port & ";" _
        & "User=" & user & ";" _
        & "Password=" & password & ";" _
        & "OPTION=3"
If cn.State = adStateOpen Then
    cn.Close
    cn.CursorLocation = adUseClient
    Set cn = New ADODB.Connection
    cn.Open StrKon
Else
cn.Open StrKon
End If
Exit Sub
buat_koneksi_Error:
    MsgBox "Ada kesalahan dengan server, periksa apakah server sudah berjalan !", vbInformation, "Cek Server"
End Sub


