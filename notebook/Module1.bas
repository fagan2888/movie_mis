Attribute VB_Name = "Module1"
Option Explicit
     Public uid As String
     Public uid1 As String
     Public mnum As String
     Public mnum1 As String
     Public cnmovie As New ADODB.Connection

Sub main()
    cnmovie.Open "driver=microsoft access driver (*.mdb);dbq=" & App.Path & "\..\database\moviemis.mdb"
    shouye.Show
End Sub

