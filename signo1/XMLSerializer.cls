VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XmlSerializer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_HelpID = 2002
Attribute VB_Description = "ComSerialization.XmlSerializer"
'--------------------------------------------------------------------------------
' Component  : XmlSerializer
' Created by : Paul Welter
' Date-Time  : 1/16/2003 - 3:01:14 PM
' ##MODULE_SUMMARY Serializes and deserializes objects into and from XML documents
'--------------------------------------------------------------------------------
' Requires Reference to
' - Microsoft XML, v3.0 (msxml3.dll)
' - TypeLib Information (Tlbinf32.dll)
'--------------------------------------------------------------------------------

Option Explicit

'##SUMMARY SAX content handler interface used for deserializing
Implements IVBSAXContentHandler

'##SUMMARY SAX error handler interface used for deserializing
Implements IVBSAXErrorHandler


'##SUMMARY SAX xml writer used for serializing
Private m_oSaxWriter As MSXML2.MXXMLWriter60

'##SUMMARY SAX content handler object used for serializing
Private m_oSaxContent As MSXML2.IVBSAXContentHandler

'##SUMMARY Type Library object used to load members of an object to serialize it
Private m_oTLI As TLi.TLIApplication

'##SUMMARY list of ignored members
Private m_strIgnoreList As String

'##SUMMARY collection of objects to deserialize to, last one is always the one written to
Private m_cObjects As Collection

'##SUMMARY the current object name that is being deserialized
Private m_sCurrentObject As String


'##SUMMARY sax reader object
Private m_oSaxReader As MSXML2.SAXXMLReader60

'##SUMMARY current sax node, used to remember the property name to write to when deserializing
Private m_sCurrentNode As String

'##SUMMARY the name of the method to call to get an object reference when deserializing a collection
Private m_strCollectionAddMethod As String

'##SUMMARY The name of the collection method that will return a reference to a collection item
Public Property Get CollectionAddMethod() As String
Attribute CollectionAddMethod.VB_Description = "The name of the collection method that will return a reference to a collection item  "
Attribute CollectionAddMethod.VB_HelpID = 2003
    CollectionAddMethod = m_strCollectionAddMethod
End Property
Public Property Let CollectionAddMethod(ByVal v_strCollectionAddMethod As String)
    m_strCollectionAddMethod = v_strCollectionAddMethod
End Property

'##SUMMARY The description of the last error that occurred
Public Property Get LastError() As String
Attribute LastError.VB_Description = "The description of the last error that occurred"
Attribute LastError.VB_HelpID = 2004
    LastError = g_strLastError
End Property

'##SUMMARY Semicolon delimited list of properties to ignore when serializing
Public Property Get IgnoreList() As String
Attribute IgnoreList.VB_Description = "Semicolon delimited list of properties to ignore when serializing  "
Attribute IgnoreList.VB_HelpID = 2005
    IgnoreList = m_strIgnoreList
End Property

Public Property Let IgnoreList(ByVal v_strIgnoreList As String)
    m_strIgnoreList = Replace(v_strIgnoreList, " ", "")    'remove spaces
End Property

Private Sub Class_Initialize()
    CallStack.ClearError    'clears any left over errors
    m_strCollectionAddMethod = "Add"
    m_strIgnoreList = "Parent"
End Sub

Private Sub IVBSAXContentHandler_characters( _
        strChars As String _
      )
Attribute IVBSAXContentHandler_characters.VB_Description = "This event is fired when the sax parser reaches the contents of an xml node"
'--------------------------------------------------------------------------------
' Procedure  : IVBSAXContentHandler_characters
' Created by : Paul Welter
' Date-Time  : 1/16/2003 - 3:02:17 PM
' ##SUMMARY This event is fired when the sax parser reaches the contents of an xml node
'--------------------------------------------------------------------------------

'<EhHeader>
    On Error GoTo IVBSAXContentHandler_characters_Error
    Dim blnErrorResume As Boolean    'tells error handler to resume after error
    Dim strErrorDescription As String    'string to set extended error information
    'adding procedure to debug call stack
    CallStack.StackAdd "XmlSerializer", "IVBSAXContentHandler_characters", strChars
    '</EhHeader>

    'set debug message
    strErrorDescription = "CurrentObject: " & m_sCurrentObject

    'if not object node and there are Chars
    If m_sCurrentObject <> m_sCurrentNode And Trim$(strChars) <> "" And m_sCurrentNode <> "" Then
        'set the properties value = the passed in chars
        LetProperty m_cObjects.item(m_cObjects.count), m_sCurrentNode, strChars
    End If

    '<EhFooter>
ExitPoint:
    CallStack.StackRemove    'this method is complete, remove from call stack

    Exit Sub
IVBSAXContentHandler_characters_Error:

    CallStack.RuntimeError "XmlSerializer", "IVBSAXContentHandler_characters", Err, Erl, strErrorDescription    'logging error
    Err.Clear
    If blnErrorResume Then Resume Next Else Resume ExitPoint
    '</EhFooter>
End Sub

Private Sub IVBSAXContentHandler_endDocument()
'do nothing
End Sub

Private Sub IVBSAXContentHandler_endElement( _
        strNamespaceURI As String, _
        strLocalName As String, _
        strQName As String _
      )
Attribute IVBSAXContentHandler_endElement.VB_Description = "This event is fired when parser reached the end of an XML node"
'--------------------------------------------------------------------------------
' Procedure  : IVBSAXContentHandler_endElement
' Created by : Paul Welter
' Date-Time  : 1/16/2003 - 3:02:57 PM
' ##SUMMARY This event is fired when parser reached the end of an XML node
'--------------------------------------------------------------------------------

'<EhHeader>
    On Error GoTo IVBSAXContentHandler_endElement_Error
    Dim blnErrorResume As Boolean    'tells error handler to resume after error
    Dim strErrorDescription As String    'string to set extended error information
    'adding procedure to debug call stack
    CallStack.StackAdd "XmlSerializer", "IVBSAXContentHandler_endElement", _
                       strNamespaceURI, strLocalName, strQName
    '</EhHeader>

    'setting debug message
    strErrorDescription = "CurrentObject: " & m_sCurrentObject

    'setting current node to empty
    m_sCurrentNode = ""

    'if the end of an object
    If strLocalName = m_sCurrentObject Then
        m_cObjects.remove m_cObjects.count    'removes last object
        If m_cObjects.count > 0 Then
            'setting current object name
            m_sCurrentObject = TypeName(m_cObjects.item(m_cObjects.count))
        End If
    End If

    '<EhFooter>
ExitPoint:
    CallStack.StackRemove    'this method is complete, remove from call stack

    Exit Sub
IVBSAXContentHandler_endElement_Error:

    CallStack.RuntimeError "XmlSerializer", "IVBSAXContentHandler_endElement", Err, Erl, strErrorDescription    'logging error
    Err.Clear
    If blnErrorResume Then Resume Next Else Resume ExitPoint
    '</EhFooter>
End Sub

Private Sub IVBSAXContentHandler_endPrefixMapping( _
        strPrefix As String _
      )
'do nothing
End Sub

Private Sub IVBSAXContentHandler_ignorableWhitespace( _
        strChars As String _
      )
'do nothing
End Sub

Private Sub IVBSAXContentHandler_processingInstruction( _
        strTarget As String, _
        strData As String _
      )
'do nothing
End Sub

Private Sub IVBSAXContentHandler_skippedEntity( _
        strName As String _
      )
'do nothing
End Sub

Private Sub IVBSAXContentHandler_startDocument()
'do nothing
End Sub

Private Sub IVBSAXContentHandler_startElement( _
        strNamespaceURI As String, _
        strLocalName As String, _
        strQName As String, _
        ByVal oAttributes As MSXML2.IVBSAXAttributes _
      )
Attribute IVBSAXContentHandler_startElement.VB_Description = "This event is fired when the parser reaches the start of an XML node"
'--------------------------------------------------------------------------------
' Procedure  : IVBSAXContentHandler_startElement
' Created by : Paul Welter
' Date-Time  : 1/16/2003 - 3:03:26 PM
' ##SUMMARY This event is fired when the parser reaches the start of an XML node
'--------------------------------------------------------------------------------

'<EhHeader>
    On Error GoTo IVBSAXContentHandler_startElement_Error
    Dim blnErrorResume As Boolean    'tells error handler to resume after error
    Dim strErrorDescription As String    'string to set extended error information
    'adding procedure to debug call stack
    CallStack.StackAdd "XmlSerializer", "IVBSAXContentHandler_startElement", _
                       strNamespaceURI, strLocalName, strQName
    '</EhHeader>

    Dim sObject As String
    Dim oChild As Object
    Dim sIndex As String
    Dim lCount As Long

    'set debug message
    strErrorDescription = "CurrentObject: " & m_sCurrentObject

    'setting current node = the element name
    m_sCurrentNode = strLocalName

    'read attributes looking for Object and Index
    With oAttributes
        For lCount = 0 To .Length - 1
            Select Case .getLocalName(lCount)
            Case "Object"    'object attribute indicates this node is the start of an object
                sObject = .GetValue(lCount)
            Case "Index"    'index attribute indicates this node is part of a collection
                sIndex = .GetValue(lCount)
            End Select
        Next
    End With

    'if is object and node not already current object
    If strLocalName <> m_sCurrentObject And sObject = "True" Then
        If IsNumeric(sIndex) Then
            'If object is from collection, use add method to return object reference to deserialize
            GetObjectProperty m_cObjects.item(m_cObjects.count), m_strCollectionAddMethod, oChild, VbMethod
        Else
            'If object is from property, call property to get object reference
            GetObjectProperty m_cObjects.item(m_cObjects.count), m_sCurrentNode, oChild
        End If

        'if we got an object, set current object name and add it to the objects collection
        If Not oChild Is Nothing Then
            m_sCurrentObject = TypeName(oChild)
            m_cObjects.Add oChild
        End If
    End If

    '<EhFooter>
ExitPoint:
    CallStack.StackRemove    'this method is complete, remove from call stack

    Exit Sub
IVBSAXContentHandler_startElement_Error:

    CallStack.RuntimeError "XmlSerializer", "IVBSAXContentHandler_startElement", Err, Erl, strErrorDescription    'logging error
    Err.Clear
    If blnErrorResume Then Resume Next Else Resume ExitPoint
    '</EhFooter>
End Sub

Private Sub IVBSAXContentHandler_startPrefixMapping( _
        strPrefix As String, _
        strURI As String _
      )
'do nothing
End Sub

Private Property Set IVBSAXContentHandler_documentLocator( _
        ByVal RHS As MSXML2.IVBSAXLocator _
      )
'do nothing
End Property

Private Sub IVBSAXErrorHandler_error( _
        ByVal oLocator As MSXML2.IVBSAXLocator, _
        strErrorMessage As String, _
        ByVal nErrorCode As Long _
      )
Attribute IVBSAXErrorHandler_error.VB_Description = "This events is called when there is a parser error"
'--------------------------------------------------------------------------------
' Procedure  : IVBSAXErrorHandler_error
' Created by : Paul Welter
' Date-Time  : 1/16/2003 - 3:03:42 PM
' ##SUMMARY This events is called when there is a parser error
'--------------------------------------------------------------------------------

'<EhHeader>
    On Error GoTo IVBSAXErrorHandler_error_Error
    Dim blnErrorResume As Boolean    'tells error handler to resume after error
    Dim strErrorDescription As String    'string to set extended error information
    'adding procedure to debug call stack
    CallStack.StackAdd "XmlSerializer", "IVBSAXErrorHandler_error"
    '</EhHeader>

    'log and raise error
    strErrorDescription = "CurrentObject: " & m_sCurrentObject
    Err.Raise nErrorCode, "SAXXMLReader", strErrorMessage

    '<EhFooter>
ExitPoint:
    CallStack.StackRemove    'this method is complete, remove from call stack

    Exit Sub
IVBSAXErrorHandler_error_Error:

    CallStack.RuntimeError "XmlSerializer", "IVBSAXErrorHandler_error", Err, Erl, strErrorDescription    'logging error
    Err.Clear
    If blnErrorResume Then Resume Next Else Resume ExitPoint
    '</EhFooter>
End Sub

Private Sub IVBSAXErrorHandler_fatalError( _
        ByVal oLocator As MSXML2.IVBSAXLocator, _
        strErrorMessage As String, _
        ByVal nErrorCode As Long _
      )
Attribute IVBSAXErrorHandler_fatalError.VB_Description = "This events is called when there is a fatal parser error"
'--------------------------------------------------------------------------------
' Procedure  : IVBSAXErrorHandler_fatalError
' Created by : Paul Welter
' Date-Time  : 1/16/2003 - 3:03:51 PM
' ##SUMMARY This events is called when there is a fatal parser error
'--------------------------------------------------------------------------------

'<EhHeader>
    On Error GoTo IVBSAXErrorHandler_fatalError_Error
    Dim blnErrorResume As Boolean    'tells error handler to resume after error
    Dim strErrorDescription As String    'string to set extended error information
    'adding procedure to debug call stack
    CallStack.StackAdd "XmlSerializer", "IVBSAXErrorHandler_fatalError"
    '</EhHeader>

    'log and raise error
    strErrorDescription = "CurrentObject: " & m_sCurrentObject
    Err.Raise nErrorCode, "SAXXMLReader", strErrorMessage

    '<EhFooter>
ExitPoint:
    CallStack.StackRemove    'this method is complete, remove from call stack

    Exit Sub
IVBSAXErrorHandler_fatalError_Error:

    CallStack.RuntimeError "XmlSerializer", "IVBSAXErrorHandler_fatalError", Err, Erl, strErrorDescription    'logging error
    Err.Clear
    If blnErrorResume Then Resume Next Else Resume ExitPoint
    '</EhFooter>
End Sub

Private Sub IVBSAXErrorHandler_ignorableWarning( _
        ByVal oLocator As MSXML2.IVBSAXLocator, _
        strErrorMessage As String, _
        ByVal nErrorCode As Long _
      )
Attribute IVBSAXErrorHandler_ignorableWarning.VB_Description = "This event is called when there is a parser warning"
'--------------------------------------------------------------------------------
' Procedure  : IVBSAXErrorHandler_ignorableWarning
' Created by : Paul Welter
' Date-Time  : 1/16/2003 - 3:19:12 PM
' ##SUMMARY This event is called when there is a parser warning
'--------------------------------------------------------------------------------

'<EhHeader>
    On Error GoTo IVBSAXErrorHandler_ignorableWarning_Error
    Dim blnErrorResume As Boolean    'tells error handler to resume after error
    Dim strErrorDescription As String    'string to set extended error information
    'adding procedure to debug call stack
    CallStack.StackAdd "XmlSerializer", "IVBSAXErrorHandler_ignorableWarning"
    '</EhHeader>

    'log and raise error
    strErrorDescription = "CurrentObject: " & m_sCurrentObject
    Err.Raise nErrorCode, "SAXXMLReader", strErrorMessage

    '<EhFooter>
ExitPoint:
    CallStack.StackRemove    'this method is complete, remove from call stack

    Exit Sub
IVBSAXErrorHandler_ignorableWarning_Error:

    CallStack.RuntimeError "XmlSerializer", "IVBSAXErrorHandler_ignorableWarning", Err, Erl, strErrorDescription    'logging error
    Err.Clear
    If blnErrorResume Then Resume Next Else Resume ExitPoint
    '</EhFooter>
End Sub

Private Function GetObjectProperty( _
        ByVal oSourceObject As Object, _
        ByVal sName As String, _
        ByRef oReturnObject As Object, _
        Optional ByVal tCallType As VbCallType = VbGet _
      ) As Boolean
Attribute GetObjectProperty.VB_Description = "wrapper for CallByName for getting an object reference"
'--------------------------------------------------------------------------------
' Procedure  : GetObjectProperty
' Created by : Paul Welter
' Date-Time  : 11/18/2002 - 12:24:18 AM
' ##SUMMARY wrapper for CallByName for getting an object reference
'--------------------------------------------------------------------------------

    On Error GoTo ErrorExit

    Set oReturnObject = CallByName(oSourceObject, sName, tCallType)

    GetObjectProperty = True

    Exit Function
ErrorExit:
    GetObjectProperty = False
End Function

Private Function GetProperty( _
        ByVal oSourceObject As Object, _
        ByVal sName As String, _
        ByRef vValue As Variant _
      ) As Boolean
Attribute GetProperty.VB_Description = "wrapper for CallByName for getting the contents of a property"
'--------------------------------------------------------------------------------
' Procedure  : GetProperty
' Created by : Paul Welter
' Date-Time  : 11/18/2002 - 12:22:41 AM
' ##SUMMARY wrapper for CallByName for getting the contents of a property
'--------------------------------------------------------------------------------
    On Error GoTo ErrorExit

    vValue = CallByName(oSourceObject, sName, VbGet)

    GetProperty = True

    Exit Function
ErrorExit:
    GetProperty = False
End Function

Private Function IsIgnored(ByVal sName As String) As Boolean
Attribute IsIgnored.VB_Description = "Checks to see if name is in ignore list"
'--------------------------------------------------------------------------------
' Procedure  : IsIgnored
' Created by : Paul Welter
' Date-Time  : 1/16/2003 - 3:19:59 PM
' ##SUMMARY Checks to see if name is in ignore list
'--------------------------------------------------------------------------------
    On Error Resume Next
    IsIgnored = InStr(1, ";" & m_strIgnoreList & ";", ";" & sName & ";")
End Function

Private Function LetProperty( _
        ByVal oSourceObject As Object, _
        ByVal sName As String, _
        ByRef vValue As Variant _
      ) As Boolean
Attribute LetProperty.VB_Description = "wrapper for CallByName for passing in the contents of a property"
'--------------------------------------------------------------------------------
' Procedure  : LetProperty
' Created by : Paul Welter
' Date-Time  : 11/18/2002 - 12:23:28 AM
' ##SUMMARY wrapper for CallByName for passing in the contents of a property
'--------------------------------------------------------------------------------

    On Error GoTo ErrorExit

    CallByName oSourceObject, sName, VbLet, vValue

    LetProperty = True

    Exit Function
ErrorExit:
    LetProperty = False
    Debug.Print Err.Description
End Function

Private Function SerializeObject( _
        ByRef oSourceObject As Object, _
        Optional ByVal lItemIndex As Long = 0, _
        Optional ByVal colname As String = "") As Boolean
'--------------------------------------------------------------------------------
' Procedure  : SerializeObject
' Created by : Paul Welter
' Date-Time  : 1/14/2003 - 4:48:22 PM
' ##SUMMARY recursive loop through objects to serialize them to xml
'--------------------------------------------------------------------------------

'<EhHeader>
    On Error GoTo SerializeObject_Error
    Dim blnErrorResume As Boolean    'tells error handler to resume after error
    Dim strErrorDescription As String    'string to set extended error information
    'adding procedure to debug call stack
    CallStack.StackAdd "XmlSerializer", "SerializeObject"
    '</EhHeader>

    Dim oInterface As TLi.InterfaceInfo
    Dim sObjectName As String

    'loading type info from object
    Set oInterface = m_oTLI.InterfaceInfoFromObject(oSourceObject)

    'get object name
    sObjectName = TypeName(oSourceObject)
    If sObjectName = "Collection" Then sObjectName = colname



    Dim oAttribute As MSXML2.SAXAttributes60
    Set oAttribute = New MSXML2.SAXAttributes60

    'add object attribute so deserializer knows this node is an object
    'oAttribute.addAttribute "", "Object", "Object", "", "True"

    'if there is an item index, this object is part of a collection
    'If lItemIndex > 0 Then
    '    oAttribute.addAttribute "", "Index", "Index", "", lItemIndex
    'End If
    m_oSaxContent.startElement "", sObjectName, sObjectName, oAttribute    'start object node

    Dim oMember As TLi.MemberInfo
    Dim sMemberName As String
    Dim oChild As Object
    Dim vValue As Variant
    Dim lCount As Long

    'loop through properties and methods
    For Each oMember In oInterface.Members
        sMemberName = oMember.Name
        'detecting collection by looking for NewEnum and making sure its a hidden member
        If (LCase$(sMemberName) = "newenum" Or LCase$(sMemberName) = "_newenum") _
           And oMember.MemberId = -4 Then
            'serialize objects in collection
            lCount = 0
            For Each oChild In oSourceObject
                lCount = lCount + 1

                SerializeObject oChild, lCount  'recursive call to serialize



            Next

            Set oChild = Nothing
        End If

        'if readable property and it doesn't have any parameters and it is not ignored
        If oMember.InvokeKind = VbGet And oMember.Parameters.count = 0 _
           And Not IsIgnored(sMemberName) Then

            'Note: All properties are treated as variant.
            'First try to read, if that fails, try reading as object
            If GetProperty(oSourceObject, sMemberName, vValue) Then
                'arrays and user defined types not supported
                If Not IsArray(vValue) And VarType(vValue) <> vbUserDefinedType _
                   And Not IsObject(vValue) And Not IsNull(vValue) Then

                    'writing xml node
                    With m_oSaxContent
                        blnErrorResume = True    'resume on error
                        .startElement "", sMemberName, sMemberName, Nothing
                        If vValue <> "" Then .Characters CStr(vValue)
                        .endElement "", sMemberName, sMemberName
                        blnErrorResume = False    'don't resume on error
                    End With

                End If
            ElseIf GetObjectProperty(oSourceObject, sMemberName, oChild) Then
                If Not oChild Is Nothing Then
                    SerializeObject oChild, , sMemberName    'recursive call to serialize
                End If
            End If

        End If
    Next

    '<EhFooter>
ExitPoint:
    'make sure we always end node
    m_oSaxContent.endElement "", sObjectName, sObjectName    'end object node

    CallStack.StackRemove    'this method is complete, remove from call stack

    Exit Function
SerializeObject_Error:

    CallStack.RuntimeError "XmlSerializer", "SerializeObject", Err, Erl, strErrorDescription    'logging error
    Err.Clear
    If blnErrorResume Then Resume Next Else Resume ExitPoint
    '</EhFooter>
End Function

Public Function Deserialize( _
       TargetObject As Object, _
       ByVal sXML As String _
     ) As Boolean
Attribute Deserialize.VB_Description = "Deserializes an XML document to the object reference"
Attribute Deserialize.VB_HelpID = 2006
'--------------------------------------------------------------------------------
' Procedure  : Deserialize
' Created by : Paul Welter
' Date-Time  : 1/15/2003 - 4:09:38 PM
' ##SUMMARY Deserializes an XML document to the object reference
'--------------------------------------------------------------------------------

'<EhHeader>
    On Error GoTo Deserialize_Error
    Dim blnErrorResume As Boolean    'tells error handler to resume after error
    Dim strErrorDescription As String    'string to set extended error information
    'adding procedure to debug call stack
    CallStack.StackAdd "XmlSerializer", "Deserialize"
    '</EhHeader>

    'creating a collection to hold objects
    Set m_cObjects = New Collection

    'getting objects name and adding it to the objects collection
    m_sCurrentObject = TypeName(TargetObject)
    m_cObjects.Add TargetObject

    'setting up the sax xml reader
    Set m_oSaxReader = New SAXXMLReader
    With m_oSaxReader
        Set .contentHandler = Me
        Set .errorhandler = Me
        .Parse sXML    'when this is called, the IVBSAXContentHandler events will fire
    End With

    Set m_oSaxReader = Nothing
    Set m_cObjects = Nothing

    Deserialize = True
    '<EhFooter>
ExitPoint:
    CallStack.StackRemove    'this method is complete, remove from call stack

    Exit Function
Deserialize_Error:

    CallStack.RuntimeError "XmlSerializer", "Deserialize", Err, Erl, strErrorDescription    'logging error
    Err.Clear
    If blnErrorResume Then Resume Next Else Resume ExitPoint
    '</EhFooter>
End Function

Public Function Serialize( _
       ByRef SourceObject As Object _
     ) As String
Attribute Serialize.VB_Description = "Serializes an object into an XML document"
Attribute Serialize.VB_HelpID = 2007

'--------------------------------------------------------------------------------
' Procedure  : Serialize
' Created by : Paul Welter
' Date-Time  : 1/14/2003 - 4:14:23 PM
' ##SUMMARY Serializes an object into an XML document
'--------------------------------------------------------------------------------

'<EhHeader>
    On Error GoTo Serialize_Error
    Dim blnErrorResume As Boolean    'tells error handler to resume after error
    Dim strErrorDescription As String    'string to set extended error information
    'adding procedure to debug call stack
    CallStack.StackAdd "XmlSerializer", "Serialize"
    '</EhHeader>

    'creating Type Library and xml writer objects
    Set m_oTLI = New TLi.TLIApplication
    Set m_oSaxWriter = New MSXML2.MXXMLWriter60
    Set m_oSaxContent = m_oSaxWriter

    'setting up sax xml writer
    With m_oSaxWriter
        .indent = True
        .encoding = "UTF-8"
        .omitXMLDeclaration = True
    End With

    m_oSaxContent.startDocument    'starting xml document
    SerializeObject SourceObject    'begin serialization
    m_oSaxContent.endDocument    'end xml doc

    Serialize = m_oSaxWriter.output    'return XML

    Set m_oSaxContent = Nothing
    Set m_oSaxWriter = Nothing
    Set m_oTLI = Nothing

    '<EhFooter>
ExitPoint:
    CallStack.StackRemove    'this method is complete, remove from call stack

    Exit Function
Serialize_Error:

    CallStack.RuntimeError "XmlSerializer", "Serialize", Err, Erl, strErrorDescription    'logging error
    Err.Clear
    If blnErrorResume Then Resume Next Else Resume ExitPoint
    '</EhFooter>
End Function

