VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***********************************************
'*
'*  This class was written by Robert May, 6/1/01
'*
'***********************************************
Option Explicit

Private m_sErrorMessage As String
Private m_lErrorNumber As Long
Private m_sErrorFilePath As String
Private m_sPosition As String
Private m_oCallStackNode As IXMLDOMNode
Private m_sAddInfo As String

Private Const CALLSTACK As String = "CallStack"
Private Const ERROROBJECT As String = "ErrorObject"

Private m_oDOM As MSXML2.DOMDocument

Public Function AddProcedureName(ByVal p_sProcedureName As String)
    Dim lRet As Long
    
    'create the node
    lRet = CreateNode(m_oCallStackNode, p_sProcedureName)
End Function
Public Function RemoveProcedureName(ByVal p_sProcedureName As String)
    Dim lRet As Long
    Dim oNode As IXMLDOMNode
    
    'locate the node
    If p_sProcedureName <> "" Then
        Set oNode = m_oCallStackNode.selectSingleNode(p_sProcedureName)
    End If
    
    'remove the node
    If Not oNode Is Nothing Then
        m_oCallStackNode.removeChild oNode
    Else
        HandleError -1, "Error Object--Procedure name """ & p_sProcedureName & """ not found."
    End If
    
    'clean up
    Set oNode = Nothing
End Function
Public Property Let Position(ByVal p_sPosition As String)
    m_sPosition = p_sPosition
End Property
Public Function HandleError(Optional ByVal p_lErrorNumber As Long, Optional ByVal p_sErrorMessage As String, Optional ByVal p_sAddInfo As String, Optional ByVal p_sErrorFileLocation As String, Optional ByVal p_bDisplayMessage As Boolean = False)
    'this routine writes out any errors that are found
    Dim lRet As Long
    Dim lfhdl As Long
    
    'set our error values
    If p_lErrorNumber <> 0 Then
        m_lErrorNumber = p_lErrorNumber
    End If
    
    'set the error messages
    If p_sErrorMessage <> "" Then
        m_sErrorMessage = p_sErrorMessage
    End If
    
    'set the additional information
    If p_sAddInfo <> "" Then
        m_sAddInfo = p_sAddInfo
    End If
    
    'set the file location
    If p_sErrorFileLocation <> "" Then
        m_sErrorFilePath = p_sErrorFileLocation
    End If
    'get a free file
    lfhdl = FreeFile
    
    'open the error location
    Open m_sErrorFilePath For Append As #lfhdl
    
    'write out the information in a parseable format
    Write #lfhdl, Now, App.EXEName, m_lErrorNumber, m_sErrorMessage, m_oDOM.xml, m_sAddInfo, m_sPosition
    
    Close lfhdl
    
    'if they want an error dialog box then
    If p_bDisplayMessage Then
        Call MsgBox("An error occured.  The error number was """ & p_lErrorNumber & """.  The error description was """ & p_sErrorMessage & """." & vbCrLf & "Error saved to """ & m_sErrorFilePath & """.", vbCritical + vbOKOnly, "Error Occured")
    End If
    
End Function
Public Property Let ErrorFilePath(ByVal p_sErrorFilePath As String)
    m_sErrorFilePath = p_sErrorFilePath
End Property
Public Property Get ErrorFilePath() As String
    ErrorFilePath = m_sErrorFilePath
End Property
Public Property Let ErrorNumber(ByVal p_lErrorNumber As Long)
    m_lErrorNumber = p_lErrorNumber
End Property
Public Property Let ErrorMessage(ByVal p_sErrorMessage As String)
    m_sErrorMessage = p_sErrorMessage
End Property
Public Property Let Information(ByVal p_sAddInfo As String)
    m_sAddInfo = p_sAddInfo
End Property

Private Sub Class_Initialize()
    Dim lRet As Long
    
    'init the dom
    Set m_oDOM = New DOMDocument
    m_oDOM.setProperty "SelectionLanguage", "XPath"
    m_oDOM.async = False
    
    'build the root xml document
    lRet = BuildRootXML()
    
    'init the error path
    m_sErrorFilePath = App.Path & "\" & App.EXEName & ".log"
End Sub

Private Sub Class_Terminate()
    Set m_oDOM = Nothing
    Set m_oCallStackNode = Nothing
    
End Sub
Private Function BuildRootXML() As Long
    Dim lRet As Long
    Dim oNode As IXMLDOMNode
    
    'create the error node
    lRet = CreateNode(m_oDOM, ERROROBJECT, oNode)
    
    'create the call stack node
    lRet = CreateNode(oNode, CALLSTACK, m_oCallStackNode)
    
    'all done
    Set oNode = Nothing
End Function
Private Function CreateNode(ByVal p_oParent As IXMLDOMNode, ByVal p_sNodeName As String, Optional ByRef r_oNewNode As IXMLDOMNode) As Long
    Dim lRet As Long
    Dim oNode As IXMLDOMNode
    
    'creat the node
    Set oNode = m_oDOM.createElement(p_sNodeName)
    
    'append the node
    Set oNode = p_oParent.appendChild(oNode)
    
    'return the node
    Set r_oNewNode = oNode
    
    'clean up
    Set oNode = Nothing
End Function
Public Property Get LastErrorNumber() As Long
    LastErrorNumber = m_lErrorNumber
End Property
Public Property Get LastErrorDescription() As String
    LastErrorDescription = m_sErrorMessage
End Property
