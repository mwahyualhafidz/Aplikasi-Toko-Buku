Attribute VB_Name = "Module1"
Public con As New ADODB.Connection
Public tabel As New ADODB.Recordset
Public sql As String

Sub bukakoneksi()
    Set con = New ADODB.Connection
    If con.State = 1 Then con.Close
        con.Open "DSN=dsbuku"
End Sub

Public Function derektoriExe() As String
    If Right(App.Path, 1) = "\" Then
        direktoriExe = App.Path
    Else
        direktoriExe = App.Path & "\"
    End If
End Function

Sub Main()
    bukakoneksi
    Form1.Show
End Sub

'Option Explicit

'Public Function JalankanSQL(SQL As String) As ADODB.Recordset
    'On Error GoTo ERR
    
    'Dim AC As New ADODB.Connection
    
    'If AC.State = adStateOpen Then AC.Close
    'Set AC = Nothing
    'AC.CursorLocation = adUseClient
    'AC.Properties.Refresh
    
    'AC.Open ("DSN=dsbuku")
    
    'Set JalankanSQL = AC.Execute(SQL)
    'Exit Function
    
'ERR:
    'MsgBox "Koneksi Ke Server ERROR!", vbCritical + vbOKOnly, "Error"
    'End
'End Function
