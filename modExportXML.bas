Attribute VB_Name = "modExportXML"
Public Sub exportCap0_12()
    Dim Dom As MSXML2.DOMDocument
    Set Dom = New MSXML2.DOMDocument
    Dom.async = False
    Dom.Load "CAP01.xml"
    
    Dim sel_node As IXMLDOMNode
    
    'GUID
    Set sel_node = Dom.selectSingleNode("DOCUMENT_RAN/HEADER/codXml")
    sel_node.Attributes(0).Text = CreateGUID
    
    'dataExport
    Set sel_node = Dom.selectSingleNode("DOCUMENT_RAN/HEADER/dataExport")
    sel_node.Text = getDateTimeXML
    
    'sirutaUAT
    
    
    Debug.Print Dom.xml
    Dom.save (App.Path & "\result.xml")
End Sub
