Attribute VB_Name = "modExportXML"
Option Explicit
Public DB As Database


Public Function Connect(strCaleDB As String) As Boolean
    On Error GoTo hErr
    If strCaleDB <> "" And FileExists(strCaleDB) = True Then
        Set DB = OpenDatabase(strCaleDB, False, False, ";UID=Admin;PWD=;")
        Connect = True
    Else
        Connect = False
    End If
    Exit Function
hErr:
    Connect = False
End Function
Public Function CloseConnection() As Boolean
    On Error GoTo hErr
    
    DB.Close
    Set DB = Nothing
    CloseConnection = True
    
    Exit Function
hErr:
        CloseConnection = False
End Function


Public Sub exportCap0_12()
    Dim Dom As MSXML2.DOMDocument
    Set Dom = New MSXML2.DOMDocument
    Dom.async = False
    Dom.Load App.Path & "\CAP0_12.xml"
    
    Dim sel_node As IXMLDOMNode
    
    'GUID
    Set sel_node = Dom.selectSingleNode("DOCUMENT_RAN/HEADER/codXml")
    sel_node.Attributes(0).Text = CreateGUID
    
    'dataExport
    Set sel_node = Dom.selectSingleNode("DOCUMENT_RAN/HEADER/dataExport")
    sel_node.Text = getDateTimeXML
    
    'sirutaUAT
    Set sel_node = Dom.selectSingleNode("DOCUMENT_RAN/HEADER/sirutaUAT")
    sel_node.Text = readSirute(readStringsUAT().localitateUAT, readStringsUAT().judetUAT).sirutaUAT
    
    
    Debug.Print Dom.xml
    Dom.save (App.Path & "\result.xml")
End Sub

