Attribute VB_Name = "modXMLProperty"
' 29.01.2002
' Juha Toivonen
' Major Blue ltd
' Finland
'

' VB6 + XML = Object DB
'
' Addition to tutorial written by Rod Stephens VB.NET + XML = Object DB
'
' http://www.informit.com/content/index.asp?product_id={A2B9405F-6B83-4474-B034-BC613B68EB1B}
'
' Why to not to use same dynamically with VB6 + XML = Object DB
'
' You need references to XML and TypeLib Information
' Function uses CallByName - method
'
'       29.01.2002
'       Toby
'
'**** .bas code
Option Explicit
  
'********************************************************************
'... return on input xmlstring will be
'  If AsAttributes Then
'    <ObjectName property1="value1" property2="value2"/>
'  Else
'    <ObjectName>
'       <property1>value1</property1>
'       <property2>value2</property2>
'    </ObjectName>
'  End If
  
' OmitProperties is given as   "Prop1,Prop2,Prop3"
'********************************************************************
Public Property Get XMLProperties(ByVal Class As Object, Optional ByVal AsAttributes = True, Optional ByVal OmitProperties As String = "XMLProperties") As String
    
    XMLProperties = myXMLProperties(Class, , AsAttributes, OmitProperties)

End Property
Public Property Let XMLProperties(ByVal Class As Object, Optional ByVal AsAttributes = True, Optional ByVal OmitProperties As String = "XMLProperties", ByVal XMLString As String)
    
    Call myXMLProperties(Class, XMLString, AsAttributes, OmitProperties)

End Property

Private Function myXMLProperties(ByVal Object As Object, _
    Optional XMLString As String = "", Optional ByVal _
    AsAttributes = True, Optional ByVal OmitProperties As _
    String = "XMLProperties") As String
  Dim tTLI      As TLIApplication
  Dim tMem      As MemberInfo
  Dim tDom      As DOMDocument
  Dim tNode     As IXMLDOMNode
  Dim tInvoke   As InvokeKinds
  Dim tOmit     As String
  Dim tName     As String     'used as lower case....
  Dim tString   As String

  Set tTLI = New TLIApplication
  Set tDom = New DOMDocument
  
  If Len(XMLString) Then
  
  '... if string given, then we are letting new property
  ' values from xmlstring
      tInvoke = VbLet
      tDom.loadXML (XMLString)
  
  Else
  
  '... else we are getting existing property values
      tInvoke = VbGet
      tDom.appendChild tDom.createNode(NODE_ELEMENT, _
          TypeName(Object), "")
  
  End If
  
  tOmit = "," & LCase(OmitProperties) & ","
  
  '... handle each get or let member from object
  
  For Each tMem In _
      TLi.InterfaceInfoFromObject(Object).Members
      
      tName = LCase(tMem.Name)
      
'      Debug.Print tName, tMem.InvokeKind
      
      '... get or let and not omitted property
      '    for example object etc...
      
      If tMem.InvokeKind = tInvoke And InStr(tOmit, "," & _
          tName & ",") = 0 _
         And tMem.Parameters.count = 0 Then
           
           
           
           
           
          If tMem.ReturnType.VarType = 0 Then ' VT_DISPATCH Then

'            '.. do nothing or do it recursive
          '   myXMLProperties tMem, False
          ElseIf tMem.ReturnType.VarType = VT_ARRAY Then
'            '.. do nothing or do it somehow else
          End If
           
          On Error Resume Next  'could be object or
              ' something else that can't handle
          If tInvoke = VbGet Then
             '... put data to XML-node
             If AsAttributes Then
               Set tNode = tDom.createAttribute(tName)
               tNode.text = CallByName(Object, tMem.Name, _
                   VbGet)
               tDom.documentElement.Attributes.setNamedItem _
                   tNode
             Else
               Set tNode = tDom.createElement(tName)
               tNode.text = CallByName(Object, tMem.Name, _
                   VbGet)
               tDom.documentElement.appendChild tNode
             End If
          Else
             '... get data from XML-node
             If AsAttributes Then
               CallByName Object, tMem.Name, VbLet, _
                 tDom.documentElement.Attributes.getNamedItem(tName).text
             Else
               CallByName Object, tMem.Name, VbLet, _
                 tDom.documentElement.selectSingleNode(tName).text
             End If
          End If
          On Error GoTo 0
      End If
  Next
  
  myXMLProperties = tDom.XML

End Function
