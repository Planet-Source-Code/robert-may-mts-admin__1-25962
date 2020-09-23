Attribute VB_Name = "modMTSAdminClient"
Option Explicit

Public Const NOERRORS As Long = 0

'we have the xml string, so why not make it a constant???
Public Const XMLSTRING As String = "<Root><ServerName/><Package><Name/><ID/><Description/><IsSystem/><Authentication/><ShutdownAfter/>" & _
        "<RunForever/><SecurityEnabled/><Identity/><Password/><Activation/><Changeable/><Deleteable/><CreatedBy/>" & _
        "<Component><ProgID/><CLSID/><Transaction/><Description/><PackageID/><PackageName/><ThreadingModel/><SecurityEnabled/>" & _
        "<DLL/><IsSystem/><Interface><Name/><ID/><Description/><ProxyCLSID/><ProxyDLL/><ProxyThreadingModel/><TypeLibID/>" & _
        "<TypeLibVersion/><TypeLibLangID/><TypeLibPlatform/><TypeLibFile/><Methods><Name/><Description/></Methods>" & _
        "</Interface></Component><FileName/></Package></Root>"
        
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long


Public Function InitDOM(ByRef r_oDOM As DOMDocument30, Optional ByVal b_ValidateOnParse As Boolean = False) As Long
    Dim lRet As Long
    Dim oErr As CError
    Const PROC_NAME As String = "InitDOM"
    Set oErr = New CError
    On Error GoTo ErrorHandler
    
'    'add us to the error stack
    oErr.AddProcedureName PROC_NAME
    
    'if we have one instantiated, to create a new one
    If r_oDOM Is Nothing Then
        Set r_oDOM = New DOMDocument30
    End If
    
    'make sure that we don't return before we're done
    r_oDOM.async = False
    
    'make sure we're using the right selection language
    r_oDOM.setProperty "SelectionLanguage", "XPath"
    
    'set the parseon load prop
    r_oDOM.validateOnParse = b_ValidateOnParse
    
Exit_Point:
    'return the results
    InitDOM = lRet
    
'    'remove us from the error stack
    oErr.RemoveProcedureName PROC_NAME
    Set oErr = Nothing
    Exit Function
ErrorHandler:
    'handle the error
    lRet = Err.Number

    oErr.HandleError lRet, Err.Description
    GoTo Exit_Point
End Function

'Private Function StubFunction() As Long
'    Dim lRet As Long
'    Const PROC_NAME As String = "StubFunction"
'    On Error GoTo ErrorHandler
'    m_oErr.AddProcedureName PROC_NAME
'
'Exit_Point:
'    'return the results
'    StubFunction = lRet
'
'    'cleanup
'
'    m_oErr.RemoveProcedureName PROC_NAME
'    Exit Function
'ErrorHandler:
'
'    lRet = Err.Number
'
'    m_oErr.HandleError lRet, Err.Description
'
'    GoTo Exit_Point
'End Function
