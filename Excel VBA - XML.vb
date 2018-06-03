http://szxml.blogspot.fr/2012/04/writing-xml-in-vba.html
http://stackoverflow.com/questions/20020990/looping-through-xml-using-vba

Import 'Microsoft XML, v6.0'


Sub read_Single_Node_XML()
	dim xDoc As IXMLDOMDocument2
    Set xDoc = New MSXML2.DOMDocument60
    
    file_Path = "D:\HTML\Exercises\VBA"
    xDoc.Load file_Path & "\stars.xml"
    
    Dim list As IXMLDOMNodeList
    Set list = xDoc.SelectNodes("//stars/kendall")
    
    Dim attr As IXMLDOMAttribute
    Dim node As IXMLDOMNode
    
	'*If children
	'Dim childNode As IXMLDOMNode
    
	'Set node = list.NextNode OR...
    Set node = list.Item(0).SelectSingleNode("name")
    '*If children
	Set childNode = node.FirstChild
    
    MsgBox childNode.Text
End Sub

'
'Write XML
Sub write_XML_File()
    Dim xDoc As Object
    Set xDoc = CreateObject("MSXML2.DOMDocument")
    
    xDoc.async = False
    xDoc.validateOnParse = False
    
    file_Path = "D:\HTML\Exercises\VBA"
    xDoc.Load file_Path & "\stars.xml"
    
    Dim list As IXMLDOMNodeList
    Set list = xDoc.SelectNodes("//stars/kendall")
    
    Dim write As IXMLDOMElement
    Set write = xDoc.createElement("eyes")
    
	write.Text = "Blue"
    list.Item(0).appendChild write
    
	'Save file
    xDoc.Save file_Path & "\stars.xml"
    
    Set xDoc = Nothing
End Sub

'
'Delete a child
Sub delete_Child()
    Dim xDoc As Object
    Set xDoc = CreateObject("MSXML2.DOMDocument")
    
    xDoc.async = False
    xDoc.validateOnParse = False
    
    file_Path = "D:\HTML\Exercises\VBA"
    xDoc.Load file_Path & "\stars.xml"
    
    Dim list As IXMLDOMNodeList
    Set list = xDoc.SelectNodes("//stars/kendall")
    
    Dim delete_Node As IXMLDOMElement
    Set delete_Node = list.Item(0).SelectSingleNode("eyes")
    
    list.Item(0).RemoveChild delete_Node
    
    xDoc.Save file_Path & "\stars.xml"
    
    Set xDoc = Nothing
End Sub

'
' Change node value
Sub change_Node_Value()
    Dim xDoc As IXMLDOMDocument2
    Set xDoc = New MSXML2.DOMDocument60
    
    xDoc.Load "D:\HTML\Exercises\VBA\stars.xml"
    
    xDoc.async = False
    xDoc.validateOnParse = False
    
    Dim list As IXMLDOMNodeList
    Set list = f.SelectNodes("//stars/kendall")
    
    Dim write_v As IXMLDOMElement
    Set write_v = list.Item(0).LastChild
	
    write_v.Text = "Blue"
    list.Item(0).appendChild write_v
    
    
    f.Save "D:\HTML\Exercises\VBA\stars.xml"
	
    Set f = Nothing
End Sub

Sub get_Attribute()
    Dim n As MSXML2.IXMLDOMDocument2
    Set n = New MSXML2.DOMDocument60
        
        
    n.Load ("D:\HTML\VBA codes\init.xml")
    
    Dim l As IXMLDOMNodeList
    Set l = n.SelectNodes("//.../...")
    
    Debug.Print l.Item(0).Attributes.getNamedItem("...").Text
    
    Set n = Nothing
    Set l = Nothing
    Set j = Nothing
End Sub