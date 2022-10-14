Attribute VB_Name = "modAsist"
Option Explicit

Public Function CreateGUID() As String
    CreateGUID = Mid$(CreateObject("Scriptlet.TypeLib").Guid, 2, 36)
End Function

Public Function getDateTimeXML() As String
     getDateTimeXML = Format(Year(Date), "0000") & "-" & Format(Month(Date), "00") & "-" & Format(Day(Date), "00") & "T" & Format(Hour(Time), "00") & ":" & Format(Minute(Time), "00") & ":" & Format(Second(Time), "00")
End Function
