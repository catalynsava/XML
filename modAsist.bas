Attribute VB_Name = "modAsist"
Option Explicit
Public Type sirute
    sirutaUAT As String
    sirutaSuper As String
    sirutaJudet As String
End Type

Public Type stringsUAT
    localitateUAT As String
    judetUAT As String
End Type

Public Function CreateGUID() As String
    CreateGUID = Mid$(CreateObject("Scriptlet.TypeLib").Guid, 2, 36)
End Function

Public Function getDateTimeXML() As String
     getDateTimeXML = Format(Year(Date), "0000") & "-" & Format(Month(Date), "00") & "-" & Format(Day(Date), "00") & "T" & Format(Hour(Time), "00") & ":" & Format(Minute(Time), "00") & ":" & Format(Second(Time), "00")
End Function
Public Function readStringsUAT() As stringsUAT
     If Connect(App.Path & "\date.mdb") Then
        Dim stringsUATtemp As stringsUAT
         Dim rsTB As Recordset
        Set rsTB = DB.OpenRecordset("SELECT * FROM DATGEN")
        stringsUATtemp.judetUAT = rsTB![judet]
        stringsUATtemp.localitateUAT = rsTB![localitate]
        readStringsUAT = stringsUATtemp
        CloseConnection
     End If
End Function

Public Function readSirute(strUAT As String, strJUDET As String) As sirute
    If Connect(App.Path & "\init.mdb") Then
        Dim rsTB As Recordset
        Dim strSQL As String
        
        Dim dblTemp As Double
        
        strSQL = "SELECT nivel0.*, judete.siruta as judsiruta FROM (nivel1 INNER JOIN nivel0 ON nivel1.sirsup = nivel0.sirsup) INNER JOIN judete ON nivel0.jud = judete.nr WHERE nivel0.denumire=""" & strUAT & """ AND nivel0.jud In (SELECT nr FROM judete WHERE denumire= """ & strJUDET & """)"
        Set rsTB = DB.OpenRecordset(strSQL)
        
        Dim siruteTMP As sirute
        
         siruteTMP.sirutaJudet = rsTB![judsiruta]
         siruteTMP.sirutaSuper = rsTB![sirsup]
         siruteTMP.sirutaUAT = rsTB![siruta]
         readSirute = siruteTMP
         CloseConnection
     End If
     
End Function

Public Function FileExists(FileName As String) As Boolean
    On Error GoTo ErrCall
    FileExists = Not (Dir$(FileName) = "")
    Exit Function
ErrCall:
    Err.Clear
    Resume Next
End Function
