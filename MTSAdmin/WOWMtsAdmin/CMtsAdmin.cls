VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "CMtsAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Because of a need to adminster Winnt 4.0 boxes from windows 2000 machines, this dll was written.
'Written by Robert May
'This code is free for you to use.  If you change it, please let me know.
'written on 8/3/01
'please note that it doesn't support roles--I didn't need that functionality.  Maybe I'll add it later.
Option Explicit
Implements ObjectControl

Private m_oCatalog As MTSAdmin.Catalog
Private m_oRoot As MTSAdmin.CatalogCollection
Private m_oErr As CError
Private m_oDOM As MSXML2.DOMDocument30

'the following is the XML that get's passed in to us.  Not all nodes are required for all methods
'see the individual methods for documentation
'for a description of what each property is, see the mts admin documentation
'if you don't want to update a property on the appropriate properties, don't send in that value
'values that are read only are marked as such, so don't bother passing them in. :)  The majority
'of the properties are at the package level
'if you want to install mutiple dlls at a given time, put them in separate component dll nodes
'<Root>
'    <ServerName/>
'    <Package>
'        <Name/>
'        <ID/> Read Only
'        <Description/>
'        <IsSystem/> Read Only
'        <Authentication/>
'        <ShutdownAfter/>
'        <RunForever/>
'        <SecurityEnabled/>
'        <Identity/>
'        <Password/> Write Only :)
'        <Activation/>
'        <Changeable/>
'        <Deleteable/>
'        <CreatedBy/>
'        <Component>
'            <ProgID/> Read Only
'            <CLSID/> Read Only
'            <Transaction/>
'            <Description/>
'            <PackageID/> Read Only
'            <PackageName/> Read Only
'            <ThreadingModel/> Read Only
'            <SecurityEnabled/>
'            <DLL/> Read Only
'            <IsSystem/> Read Only
'            <Interface> Read Only
'                <Name/> Read Only
'                <ID/> Read Only
'                <Description/> Read Only
'                <ProxyCLSID/> Read Only
'                <ProxyDLL/> Read Only
'                <ProxyThreadingModel/> Read Only
'                <TypeLibID/> Read Only
'                <TypeLibVersion/> Read Only
'                <TypeLibLangID/> Read Only
'                <TypeLibPlatform/> Read Only
'                <TypeLibFile/> Read Only
'                <Methods> Read Only
'                    <Name/> Read Only
'                    <Description/> Read Only
'                </Methods>
'            </Interface>
'        </Component>
'    </Package>
'    <FileName/>
'</Root>

Private Const NOERRORS As Long = 0

Public Function AddPackage(ByVal p_sInput As String, ByRef r_sOutput As Variant) As Long
    Dim lRet As Long
    Dim oPackages As IXMLDOMNodeList
    Dim oMTSPackages As MTSAdmin.CatalogCollection
    Dim oPackage As MTSAdmin.CatalogObject
    Dim i As Long
    Const PROC_NAME As String = "AddPackage"
    On Error GoTo ErrorHandler
    
    'add us to the error stack
    m_oErr.AddProcedureName PROC_NAME
        
    'load the dom
    If lRet = NOERRORS Then
        lRet = m_oDOM.loadXML(p_sInput)
    End If
    
    'check the results
    If lRet Then
        lRet = NOERRORS
    Else
        Err.Raise m_oDOM.parseError.errorCode, PROC_NAME, m_oDOM.parseError.reason
    End If
    
    'get the packages
    If lRet = NOERRORS Then
        lRet = GetPackages(oMTSPackages, oPackages)
    End If
    
    'loop through the provided packages
    If lRet = NOERRORS Then
        For i = 0 To oPackages.length - 1
            'Add the package
            If lRet = NOERRORS Then
                Set oPackage = oMTSPackages.Add
            End If
            
            'set it's properties
            If lRet = NOERRORS Then
                SetValue oPackage, oPackages.Item(i), "Name"
                SetValue oPackage, oPackages.Item(i), "ID"
                SetValue oPackage, oPackages.Item(i), "Description"
                SetValue oPackage, oPackages.Item(i), "Authentication"
                SetValue oPackage, oPackages.Item(i), "ShutdownAfter"
                SetValue oPackage, oPackages.Item(i), "RunForever"
                SetValue oPackage, oPackages.Item(i), "SecurityEnabled"
                SetValue oPackage, oPackages.Item(i), "Identity"
                SetValue oPackage, oPackages.Item(i), "Password"
                SetValue oPackage, oPackages.Item(i), "Activation"
                SetValue oPackage, oPackages.Item(i), "Changeable"
                SetValue oPackage, oPackages.Item(i), "Deleteable"
                SetValue oPackage, oPackages.Item(i), "CreatedBy"
            End If
            
            'save the changes
            If lRet = NOERRORS Then
                oMTSPackages.SaveChanges
            End If
        Next i
    End If
Exit_Point:
    'return the results
    AddPackage = lRet
    
    'clean up
    Set oPackages = Nothing
    Set oMTSPackages = Nothing
    Set oPackage = Nothing
    
    
    'remove us from the error stack
    m_oErr.RemoveProcedureName PROC_NAME
    Exit Function
ErrorHandler:
    'handle the error
    lRet = Err.Number
    
    m_oErr.HandleError lRet, Err.Description
    GoTo Exit_Point

End Function
Public Function RemovePackage(ByVal p_sInput As String, ByRef r_sOutput As Variant) As Long
    Dim lRet As Long
    Dim oPackages As IXMLDOMNodeList
    Dim oMTSPackages As MTSAdmin.CatalogCollection
    Dim i As Long
    Dim lPackageID As Long
    Const PROC_NAME As String = "RemovePackage"
    On Error GoTo ErrorHandler
    
    'add us to the error stack
    m_oErr.AddProcedureName PROC_NAME
    
    'load the dom
    If lRet = NOERRORS Then
        lRet = m_oDOM.loadXML(p_sInput)
    End If
    
    'check the results
    If lRet Then
        lRet = NOERRORS
    Else
        Err.Raise m_oDOM.parseError.errorCode, PROC_NAME, m_oDOM.parseError.reason
    End If
    
    'get the packages
    If lRet = NOERRORS Then
        lRet = GetPackages(oMTSPackages, oPackages)
    End If
    
    'loop through the provided packages
    If lRet = NOERRORS Then
        For i = 0 To oPackages.length - 1
            'find the package
            If lRet = NOERRORS Then
                'get the packageid
                If lRet = NOERRORS Then
                    lRet = FindPackage(oMTSPackages, oPackages.Item(i), lPackageID)
                End If
                
                'delete the package
                If lRet = NOERRORS Then
                    oMTSPackages.Remove lPackageID
                End If
            End If
        Next i
        
        'save the changes
        If lRet = NOERRORS Then
            oMTSPackages.SaveChanges
        End If
    End If

Exit_Point:
    'return the results
    RemovePackage = lRet
    
    Set oPackages = Nothing
    Set oMTSPackages = Nothing
    
    'remove us from the error stack
    m_oErr.RemoveProcedureName PROC_NAME
    Exit Function
ErrorHandler:
    'handle the error
    lRet = Err.Number
    
    m_oErr.HandleError lRet, Err.Description
    GoTo Exit_Point

End Function
Public Function AddDLL(ByVal p_sInput As String, ByRef r_sOutput As Variant) As Long
    Dim lRet As Long
    Dim oPackages As IXMLDOMNodeList
    Dim oComponents As IXMLDOMNodeList
    Dim oMTSPackages As MTSAdmin.CatalogCollection
    Dim oMTSComponents As MTSAdmin.CatalogCollection
    Dim oComponentUtil As MTSAdmin.ComponentUtil
    Dim lComponentID As Long
    Dim lPackageID As Long
    Dim i As Long
    Dim j As Long
    Const PROC_NAME As String = "AddDLL"
    On Error GoTo ErrorHandler
    
    'add us to the error stack
    m_oErr.AddProcedureName PROC_NAME
    
    'load the dom
    If lRet = NOERRORS Then
        lRet = m_oDOM.loadXML(p_sInput)
    End If
    
    'check the results
    If lRet Then
        lRet = NOERRORS
    Else
        Err.Raise m_oDOM.parseError.errorCode, PROC_NAME, m_oDOM.parseError.reason
    End If
    
    'get the packages
    If lRet = NOERRORS Then
        lRet = GetPackages(oMTSPackages, oPackages)
    End If
    
    'loop through the provided packages
    If lRet = NOERRORS Then
        For i = 0 To oPackages.length - 1
            'find the package
            If lRet = NOERRORS Then
                'get the packageid
                If lRet = NOERRORS Then
                    lRet = FindPackage(oMTSPackages, oPackages.Item(i), lPackageID)
                End If
                
                'get the supplied components
                If lRet = NOERRORS Then
                    lRet = GetComponents(oMTSComponents, oComponents, oPackages.Item(i), oMTSPackages, oMTSPackages.Item(lPackageID).Value("ID"))
                End If
                
                'loop through the components and add them
                If lRet = NOERRORS Then
                    'get the util interface
                    Set oComponentUtil = oMTSComponents.GetUtilInterface
                    
                    'loop through the supplied components
                    For j = 0 To oComponents.length - 1
                        'add the dll
                        If lRet = NOERRORS Then
                            oComponentUtil.InstallComponent NodeValue("DLL", oComponents.Item(j)), "", ""
                        End If
                    Next j
                End If
            End If
        Next i
    End If
        
Exit_Point:
    'return the results
    AddDLL = lRet
    
    'remove us from the error stack
    m_oErr.RemoveProcedureName PROC_NAME
    Exit Function
ErrorHandler:
    'handle the error
    lRet = Err.Number
    
    m_oErr.HandleError lRet, Err.Description
    GoTo Exit_Point

End Function
Public Function RemoveComponent(ByVal p_sInput As String, ByRef r_sOutput As Variant) As Long
    Dim lRet As Long
    Dim oPackages As IXMLDOMNodeList
    Dim oComponents As IXMLDOMNodeList
    Dim oMTSPackages As MTSAdmin.CatalogCollection
    Dim oMTSComponents As MTSAdmin.CatalogCollection
    Dim lComponentID As Long
    Dim lPackageID As Long
    Dim i As Long
    Dim j As Long
    Const PROC_NAME As String = "RemoveComponent"
    On Error GoTo ErrorHandler
    
    'add us to the error stack
    m_oErr.AddProcedureName PROC_NAME
    
    'load the dom
    If lRet = NOERRORS Then
        lRet = m_oDOM.loadXML(p_sInput)
    End If
    
    'check the results
    If lRet Then
        lRet = NOERRORS
    Else
        Err.Raise m_oDOM.parseError.errorCode, PROC_NAME, m_oDOM.parseError.reason
    End If
    
    'get the packages
    If lRet = NOERRORS Then
        lRet = GetPackages(oMTSPackages, oPackages)
    End If
    
    'loop through the provided packages
    If lRet = NOERRORS Then
        For i = 0 To oPackages.length - 1
            'find the package
            If lRet = NOERRORS Then
                'get the packageid
                If lRet = NOERRORS Then
                    lRet = FindPackage(oMTSPackages, oPackages.Item(i), lPackageID)
                End If
                
                'get the supplied components
                If lRet = NOERRORS Then
                    lRet = GetComponents(oMTSComponents, oComponents, oPackages.Item(i), oMTSPackages, oMTSPackages.Item(lPackageID).Value("ID"))
                End If
                
                'loop through the components and add them
                If lRet = NOERRORS Then
                    For j = 0 To oComponents.length - 1
                        If lRet = NOERRORS Then
                            'find the component
                            lRet = FindComponent(oMTSComponents, oComponents.Item(j), lComponentID)
                        End If
                        
                        'remove the component
                        If lRet = NOERRORS Then
                            oMTSComponents.Remove lComponentID
                        End If
                    Next j
                    
                    'save the changes
                    If lRet = NOERRORS Then
                        oMTSComponents.SaveChanges
                    End If
                End If
            End If
        Next i
    End If
        
Exit_Point:
    'return the results
    RemoveComponent = lRet
    
    'clean up
    Set oPackages = Nothing
    Set oComponents = Nothing
    Set oMTSPackages = Nothing
    Set oMTSComponents = Nothing
    
    'remove us from the error stack
    m_oErr.RemoveProcedureName PROC_NAME
    Exit Function
ErrorHandler:
    'handle the error
    lRet = Err.Number
    
    m_oErr.HandleError lRet, Err.Description
    GoTo Exit_Point

End Function
Public Function ListComponents(ByVal p_sInput As String, ByRef r_sOutput As Variant) As Long
    Dim lRet As Long
    Dim oPackages As IXMLDOMNodeList
    Dim oComponents As IXMLDOMNodeList
    Dim oMTSPackages As MTSAdmin.CatalogCollection
    Dim oMTSComponents As MTSAdmin.CatalogCollection
    Dim oPropertyInfo As MTSAdmin.CatalogCollection
    Dim lComponentID As Long
    Dim lPackageID As Long
    Dim i As Long
    Dim j As Long
    Const PROC_NAME As String = "ListComponents"
    On Error GoTo ErrorHandler
    
    'add us to the error stack
    m_oErr.AddProcedureName PROC_NAME
        
    'load the dom
    If lRet = NOERRORS Then
        lRet = m_oDOM.loadXML(p_sInput)
    End If
    
    'check the results
    If lRet Then
        lRet = NOERRORS
    Else
        Err.Raise m_oDOM.parseError.errorCode, PROC_NAME, m_oDOM.parseError.reason
    End If
    
    'get the packages
    If lRet = NOERRORS Then
        lRet = GetPackages(oMTSPackages, oPackages)
    End If
    
    'loop through the provided packages
    If lRet = NOERRORS Then
        For i = 0 To oPackages.length - 1
            'find the package
            If lRet = NOERRORS Then
                'get the packageid
                If lRet = NOERRORS Then
                    lRet = FindPackage(oMTSPackages, oPackages.Item(i), lPackageID)
                End If
                
                'get the supplied components
                If lRet = NOERRORS Then
                    lRet = GetComponents(oMTSComponents, oComponents, oPackages.Item(i), oMTSPackages, oMTSPackages.Item(lPackageID).Value("ID"))
                End If
                
                'loop through the components and add them to the dom
                If lRet = NOERRORS Then
                                        
                    'loop through the components
                    For j = 0 To oMTSComponents.Count - 1
                        Set oPropertyInfo = oMTSComponents.GetCollection("PropertyInfo", oMTSComponents.Item(j).Key)
                        lRet = AddXML(oPackages.Item(i), "Component", oPropertyInfo, oMTSComponents.Item(j))
                    Next j
                End If
            End If
        Next i
    End If
    
Exit_Point:
    'return the results
    ListComponents = lRet
    r_sOutput = m_oDOM.xml
        
    'remove us from the error stack
    m_oErr.RemoveProcedureName PROC_NAME
    Exit Function
ErrorHandler:
    'handle the error
    lRet = Err.Number
    
    m_oErr.HandleError lRet, Err.Description
    GoTo Exit_Point

End Function
Public Function ShutDownPackage(ByVal p_sInput As String, ByRef r_sOutput As Variant) As Long
    Dim lRet As Long
    Dim oPackages As IXMLDOMNodeList
    Dim oMTSPackages As MTSAdmin.CatalogCollection
    Dim oPackageUtil As MTSAdmin.PackageUtil
    Dim i As Long
    Dim lPackageID As Long
    Const PROC_NAME As String = "ShutDownPackage"
    On Error GoTo ErrorHandler
    
    'add us to the error stack
    m_oErr.AddProcedureName PROC_NAME
    
    'load the dom
    If lRet = NOERRORS Then
        lRet = m_oDOM.loadXML(p_sInput)
    End If
    
    'check the results
    If lRet Then
        lRet = NOERRORS
    Else
        Err.Raise m_oDOM.parseError.errorCode, PROC_NAME, m_oDOM.parseError.reason
    End If
    
    'get the packages
    If lRet = NOERRORS Then
        lRet = GetPackages(oMTSPackages, oPackages)
    End If
    
    'loop through the provided packages
    If lRet = NOERRORS Then
        For i = 0 To oPackages.length - 1
            'find the package
            If lRet = NOERRORS Then
                'get the packageid
                If lRet = NOERRORS Then
                    lRet = FindPackage(oMTSPackages, oPackages.Item(i), lPackageID)
                End If
                
                'get the package util
                If lRet = NOERRORS Then
                    Set oPackageUtil = oMTSPackages.GetUtilInterface
                End If
                
                'shut down the package
                If lRet = NOERRORS Then
                    oPackageUtil.ShutDownPackage oMTSPackages.Item(lPackageID).Value("ID")
                End If
            End If
        Next i
        
    End If

Exit_Point:
    'return the results
    ShutDownPackage = lRet
    
    Set oPackages = Nothing
    Set oMTSPackages = Nothing
    
    'remove us from the error stack
    m_oErr.RemoveProcedureName PROC_NAME
    Exit Function
ErrorHandler:
    'handle the error
    lRet = Err.Number
    
    m_oErr.HandleError lRet, Err.Description
    GoTo Exit_Point

End Function
Public Function SetPackageProperties(ByVal p_sInput As String, ByRef r_sOutput As Variant) As Long
    Dim lRet As Long
    Dim oPackages As IXMLDOMNodeList
    Dim oMTSPackages As MTSAdmin.CatalogCollection
    Dim oPackage As MTSAdmin.CatalogObject
    Dim lPackageID As Long
    Dim i As Long
    Const PROC_NAME As String = "SetPackageProperties"
    On Error GoTo ErrorHandler
    
    'add us to the error stack
    m_oErr.AddProcedureName PROC_NAME
        
    'load the dom
    If lRet = NOERRORS Then
        lRet = m_oDOM.loadXML(p_sInput)
    End If
    
    'check the results
    If lRet Then
        lRet = NOERRORS
    Else
        Err.Raise m_oDOM.parseError.errorCode, PROC_NAME, m_oDOM.parseError.reason
    End If
    
    'get the packages
    If lRet = NOERRORS Then
        lRet = GetPackages(oMTSPackages, oPackages)
    End If
    
    'loop through the provided packages
    If lRet = NOERRORS Then
        For i = 0 To oPackages.length - 1
            'find the package
            If lRet = NOERRORS Then
                'get the packageid
                If lRet = NOERRORS Then
                    lRet = FindPackage(oMTSPackages, oPackages.Item(i), lPackageID, oPackage)
                End If
                
                'set it's property values
                If lRet = NOERRORS Then
                    lRet = SetValues(oMTSPackages, lPackageID, oPackages.Item(i))
                End If
                                
                'save the changes
                If lRet = NOERRORS Then
                    oMTSPackages.SaveChanges
                End If
            End If
        Next i
    End If
Exit_Point:
    'return the results
    SetPackageProperties = lRet
    
    'clean up
    Set oPackages = Nothing
    Set oMTSPackages = Nothing
    Set oPackage = Nothing
    
    'remove us from the error stack
    m_oErr.RemoveProcedureName PROC_NAME
    Exit Function
ErrorHandler:
    'handle the error
    lRet = Err.Number
    
    m_oErr.HandleError lRet, Err.Description
    GoTo Exit_Point

End Function
Public Function SetComponentProperties(ByVal p_sInput As String, ByRef r_sOutput As Variant) As Long
    Dim lRet As Long
    Dim oPackages As IXMLDOMNodeList
    Dim oComponents As IXMLDOMNodeList
    Dim oMTSPackages As MTSAdmin.CatalogCollection
    Dim oMTSComponents As MTSAdmin.CatalogCollection
    Dim lComponentID As Long
    Dim lPackageID As Long
    Dim i As Long
    Dim j As Long
    Const PROC_NAME As String = "SetComponentProperties"
    On Error GoTo ErrorHandler
    
    'add us to the error stack
    m_oErr.AddProcedureName PROC_NAME
    
    'load the dom
    If lRet = NOERRORS Then
        lRet = m_oDOM.loadXML(p_sInput)
    End If
    
    'check the results
    If lRet Then
        lRet = NOERRORS
    Else
        Err.Raise m_oDOM.parseError.errorCode, PROC_NAME, m_oDOM.parseError.reason
    End If
    
    'get the packages
    If lRet = NOERRORS Then
        lRet = GetPackages(oMTSPackages, oPackages)
    End If
    
    'loop through the provided packages
    If lRet = NOERRORS Then
        For i = 0 To oPackages.length - 1
            'find the package
            If lRet = NOERRORS Then
                'get the packageid
                If lRet = NOERRORS Then
                    lRet = FindPackage(oMTSPackages, oPackages.Item(i), lPackageID)
                End If
                
                'get the supplied components
                If lRet = NOERRORS Then
                    lRet = GetComponents(oMTSComponents, oComponents, oPackages.Item(i), oMTSPackages, oMTSPackages.Item(lPackageID).Value("ID"))
                End If
                
                'loop through the components and add them
                If lRet = NOERRORS Then
                    For j = 0 To oComponents.length - 1
                        If lRet = NOERRORS Then
                            'find the component
                            lRet = FindComponent(oMTSComponents, oComponents.Item(j), lComponentID)
                        End If
                        
                        'set it's property values
                        If lRet = NOERRORS Then
                            lRet = SetValues(oMTSComponents, lComponentID, oComponents.Item(j))
                        End If
                        
                    Next j
                    
                    'save the changes
                    If lRet = NOERRORS Then
                        oMTSComponents.SaveChanges
                    End If
                End If
            End If
        Next i
    End If
        
Exit_Point:
    'return the results
    SetComponentProperties = lRet
    
    'clean up
    Set oPackages = Nothing
    Set oComponents = Nothing
    Set oMTSPackages = Nothing
    Set oMTSComponents = Nothing
    
    'remove us from the error stack
    m_oErr.RemoveProcedureName PROC_NAME
    Exit Function
ErrorHandler:
    'handle the error
    lRet = Err.Number
    
    m_oErr.HandleError lRet, Err.Description
    GoTo Exit_Point

End Function
Public Function ExportPackage(ByVal p_sInput As String, ByRef r_sOutput As Variant) As Long
    Dim lRet As Long
    Dim oPackages As IXMLDOMNodeList
    Dim oMTSPackages As MTSAdmin.CatalogCollection
    Dim oPackageUtil As MTSAdmin.PackageUtil
    Dim lPackageID As Long
    Dim i As Long
    Const PROC_NAME As String = "ExportPackage"
    On Error GoTo ErrorHandler
    
    'add us to the error stack
    m_oErr.AddProcedureName PROC_NAME
    
    'load the dom
    If lRet = NOERRORS Then
        lRet = m_oDOM.loadXML(p_sInput)
    End If
    
    'check the results
    If lRet Then
        lRet = NOERRORS
    Else
        Err.Raise m_oDOM.parseError.errorCode, PROC_NAME, m_oDOM.parseError.reason
    End If
    
    'get the packages
    If lRet = NOERRORS Then
        lRet = GetPackages(oMTSPackages, oPackages)
    End If
    
    'loop through the provided packages
    If lRet = NOERRORS Then
        For i = 0 To oPackages.length - 1
            'find the package
            If lRet = NOERRORS Then
                'get the packageid
                If lRet = NOERRORS Then
                    lRet = FindPackage(oMTSPackages, oPackages.Item(i), lPackageID)
                End If
                            
                'export the package
                If lRet = NOERRORS Then
                    Set oPackageUtil = oMTSPackages.GetUtilInterface
                    oPackageUtil.ExportPackage oMTSPackages.Item(lPackageID).Value("ID"), NodeValue("FileName", oPackages.Item(i)), mtsExportUsers
                End If
            End If
        Next i
    End If
    
Exit_Point:
    'return the results
    ExportPackage = lRet
    
    'clean up
    Set oPackages = Nothing
    Set oMTSPackages = Nothing
    Set oPackageUtil = Nothing
    
    'remove us from the error stack
    m_oErr.RemoveProcedureName PROC_NAME
    Exit Function
ErrorHandler:
    'handle the error
    lRet = Err.Number
    
    m_oErr.HandleError lRet, Err.Description
    GoTo Exit_Point
End Function


Private Sub ObjectControl_Activate()
    On Error Resume Next
    Dim lRet As Long
    Const PROC_NAME As String = "ObjectControl_Activate"
    Set m_oErr = New CError
    
    On Error GoTo ErrorHandler
    
    m_oErr.AddProcedureName PROC_NAME
    
    'init our dom
    InitDOM m_oDOM
        
    'init the catalog
    Set m_oCatalog = New MTSAdmin.Catalog
Exit_Point:
    
    'remove us from the error stack
    m_oErr.RemoveProcedureName PROC_NAME
    Exit Sub
ErrorHandler:
    'handle the error
    lRet = Err.Number
    
    m_oErr.HandleError lRet, Err.Description
    GoTo Exit_Point
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function

Private Sub ObjectControl_Deactivate()
    Set m_oErr = Nothing
    Set m_oDOM = Nothing
    Set m_oCatalog = Nothing
    Set m_oRoot = Nothing
End Sub
Private Function StubFunction() As Long
    Dim lRet As Long
    Const PROC_NAME As String = "StubFunction"
    On Error GoTo ErrorHandler
    
    'add us to the error stack
    m_oErr.AddProcedureName PROC_NAME
    
Exit_Point:
    'return the results
    StubFunction = lRet
    
    'remove us from the error stack
    m_oErr.RemoveProcedureName PROC_NAME
    Exit Function
ErrorHandler:
    'handle the error
    lRet = Err.Number
    
    m_oErr.HandleError lRet, Err.Description
    GoTo Exit_Point
End Function
Private Function InitDOM(ByRef r_oDOM As DOMDocument30, Optional ByVal b_ValidateOnParse As Boolean = False) As Long
    Dim lRet As Long
    Const PROC_NAME As String = "InitDOM"
    On Error GoTo ErrorHandler
    
    'add us to the error stack
    m_oErr.AddProcedureName PROC_NAME
    
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
    
    'remove us from the error stack
    m_oErr.RemoveProcedureName PROC_NAME
    Exit Function
ErrorHandler:
    'handle the error
    lRet = Err.Number
    
    m_oErr.HandleError lRet, Err.Description
    GoTo Exit_Point
End Function

Private Function Connect(ByVal p_sComputer As String) As Long
    Dim lRet As Long
    Const PROC_NAME As String = "Connect"
    On Error GoTo ErrorHandler
    
    'add us to the error stack
    m_oErr.AddProcedureName PROC_NAME
    
    'connect to the root computer
    Set m_oRoot = m_oCatalog.Connect(p_sComputer)
    
Exit_Point:
    
    'return the results
    Connect = lRet
    
    'remove us from the error stack
    m_oErr.RemoveProcedureName PROC_NAME
    Exit Function
ErrorHandler:
    'handle the error
    lRet = Err.Number
    
    m_oErr.HandleError lRet, Err.Description
    GoTo Exit_Point
End Function

Private Function NodeValue(ByVal p_sNodeName As String, Optional ByVal p_oParent As IXMLDOMNode) As String
    Dim lRet As Long
    Dim oParent As IXMLDOMNode
    Dim sValue As String
    Dim oNode As IXMLDOMNode
    Const PROC_NAME As String = "NodeValue"
    On Error GoTo ErrorHandler
    
    'add us to the error stack
    m_oErr.AddProcedureName PROC_NAME
    
    'get the appropriate starting place
    If p_oParent Is Nothing Then
        Set oParent = m_oDOM
    Else
        Set oParent = p_oParent
    End If
    
    'look up the value
    Set oNode = oParent.selectSingleNode(p_sNodeName)
    
    'check for results
    If Not oNode Is Nothing Then
        'check to see if they have nodes
        If oNode.hasChildNodes = True Then
            'check to see if the first child is text--otherwise, they probably didn't put a value in
            'they may have put the XML in the wrong order.  Let's hope not. :)
            If oNode.firstChild.nodeType = NODE_TEXT Then
                sValue = oNode.firstChild.Text
            Else
                sValue = ""
            End If
        Else
            sValue = ""
        End If
    Else
        sValue = ""
    End If
    
Exit_Point:
    'return the results
    NodeValue = sValue
    
    'clean up
    Set oParent = Nothing
    Set oNode = Nothing
    
    'remove us from the error stack
    m_oErr.RemoveProcedureName PROC_NAME
    Exit Function
ErrorHandler:
    'handle the error
    lRet = Err.Number
    
    m_oErr.HandleError lRet, Err.Description
    GoTo Exit_Point
End Function

Private Function SetValue(ByRef r_oObject As MTSAdmin.CatalogObject, ByVal r_oNode As IXMLDOMNode, ByVal p_sValueName As String) As Long
    Dim lRet As Long
    Const PROC_NAME As String = "SetValue"
    On Error GoTo ErrorHandler
    
    'add us to the error stack
    m_oErr.AddProcedureName PROC_NAME
    
    'check to see if the node exists
    If NodeExists(p_sValueName, r_oNode) Then
        'check to make sure that we can set the property
        If Not r_oObject.IsPropertyReadOnly(p_sValueName) Then
            'add the property
            If IsNumeric(NodeValue(p_sValueName, r_oNode)) Then
                r_oObject.Value(p_sValueName) = CLng(NodeValue(p_sValueName, r_oNode))
            Else
                r_oObject.Value(p_sValueName) = NodeValue(p_sValueName, r_oNode)
            End If
        End If
            
    End If
    
Exit_Point:
    'return the results
    SetValue = lRet
    
    'remove us from the error stack
    m_oErr.RemoveProcedureName PROC_NAME
    Exit Function
ErrorHandler:
    'handle the error
    lRet = Err.Number
    
    m_oErr.HandleError lRet, Err.Description
    GoTo Exit_Point
End Function

Private Function NodeExists(ByVal p_sNodeName As String, Optional ByVal p_oParent As IXMLDOMNode) As Boolean
    Dim lRet As Long
    Dim oNode As IXMLDOMNode
    Dim oParent As IXMLDOMNode
    Dim bExists As Boolean
    Const PROC_NAME As String = "NodeExists"
    On Error GoTo ErrorHandler
    
    'add us to the error stack
    m_oErr.AddProcedureName PROC_NAME
    
    'get the right parent
    If p_oParent Is Nothing Then
        Set oParent = m_oDOM
    Else
        Set oParent = p_oParent
    End If
    
    'attempt to get the node
    Set oNode = oParent.selectSingleNode(p_sNodeName)
    
    'check to see if we found the node
    If oNode Is Nothing Then
        bExists = False
    Else
        bExists = True
    End If
    
Exit_Point:
    'return the results
    NodeExists = bExists
    
    'clean up
    Set oNode = Nothing
    Set oParent = Nothing
    
    'remove us from the error stack
    m_oErr.RemoveProcedureName PROC_NAME
    Exit Function
ErrorHandler:
    'handle the error
    lRet = Err.Number
    
    m_oErr.HandleError lRet, Err.Description
    GoTo Exit_Point
End Function

Public Function ListPackages(ByVal p_sInput As String, ByRef r_sOutput As Variant) As Long
    Dim lRet As Long
    Dim oPackages As IXMLDOMNodeList
    Dim oMTSPackages As MTSAdmin.CatalogCollection
    Dim oPropertyInfo As MTSAdmin.CatalogCollection
    Dim i As Long
    Const PROC_NAME As String = "ListPackages"
    On Error GoTo ErrorHandler
    
    'add us to the error stack
    m_oErr.AddProcedureName PROC_NAME
    
    'load the dom
    If lRet = NOERRORS Then
        lRet = m_oDOM.loadXML(p_sInput)
    End If
    
    'check the results
    If lRet Then
        lRet = NOERRORS
    Else
        Err.Raise m_oDOM.parseError.errorCode, PROC_NAME, m_oDOM.parseError.reason
    End If
    
    'get the packages
    If lRet = NOERRORS Then
        lRet = GetPackages(oMTSPackages, oPackages)
    End If
    
    'loop through the provided packages
    If lRet = NOERRORS Then
        For i = 0 To oMTSPackages.Count - 1
            Set oPropertyInfo = oMTSPackages.GetCollection("PropertyInfo", oMTSPackages.Item(i).Key)
            lRet = AddXML(m_oDOM.selectSingleNode("/Root"), "Package", oPropertyInfo, oMTSPackages.Item(i))
        Next i
    End If
    
Exit_Point:
    'return the results
    ListPackages = lRet
    r_sOutput = CStr(m_oDOM.xml)
    
    'clean up
    Set oPackages = Nothing
    Set oMTSPackages = Nothing
    
    'remove us from the error stack
    m_oErr.RemoveProcedureName PROC_NAME
    Exit Function
ErrorHandler:
    'handle the error
    lRet = Err.Number
    
    m_oErr.HandleError lRet, Err.Description
    GoTo Exit_Point
End Function
Public Function ListInterfaces(ByVal p_sInput As String, ByRef r_sOutput As Variant) As Long
    Dim lRet As Long
    Const PROC_NAME As String = "ListInterfaces"
    On Error GoTo ErrorHandler
    
    'add us to the error stack
    m_oErr.AddProcedureName PROC_NAME
    
    'load the dom
    If lRet = NOERRORS Then
        lRet = m_oDOM.loadXML(p_sInput)
    End If
    
    'check the results
    If lRet Then
        lRet = NOERRORS
    Else
        Err.Raise m_oDOM.parseError.errorCode, PROC_NAME, m_oDOM.parseError.reason
    End If
Exit_Point:
    'return the results
    ListInterfaces = lRet
    
    'remove us from the error stack
    m_oErr.RemoveProcedureName PROC_NAME
    Exit Function
ErrorHandler:
    'handle the error
    lRet = Err.Number
    
    m_oErr.HandleError lRet, Err.Description
    GoTo Exit_Point
End Function
Public Function ListMethods(ByVal p_sInput As String, ByRef r_sOutput As Variant) As Long
    Dim lRet As Long
    Const PROC_NAME As String = "ListMethods"
    On Error GoTo ErrorHandler
    
    'add us to the error stack
    m_oErr.AddProcedureName PROC_NAME
    
    'load the dom
    If lRet = NOERRORS Then
        lRet = m_oDOM.loadXML(p_sInput)
    End If
    
    'check the results
    If lRet Then
        lRet = NOERRORS
    Else
        Err.Raise m_oDOM.parseError.errorCode, PROC_NAME, m_oDOM.parseError.reason
    End If
Exit_Point:
    'return the results
    ListMethods = lRet
    
    'remove us from the error stack
    m_oErr.RemoveProcedureName PROC_NAME
    Exit Function
ErrorHandler:
    'handle the error
    lRet = Err.Number
    
    m_oErr.HandleError lRet, Err.Description
    GoTo Exit_Point
End Function

Private Function GetPackages(ByRef r_oMTSPackages As MTSAdmin.CatalogCollection, ByRef r_oPackages As IXMLDOMNodeList) As Long
    Dim lRet As Long
    Dim oPackages As IXMLDOMNodeList
    Dim oMTSPackages As MTSAdmin.CatalogCollection
    Const PROC_NAME As String = "GetPackages"
    On Error GoTo ErrorHandler
    
    'add us to the error stack
    m_oErr.AddProcedureName PROC_NAME
    
    'get all of the packages they want to add--we only allow them to do one server at a time.
    If lRet = NOERRORS Then
        Set oPackages = m_oDOM.selectNodes("//Package")
    End If
    
    'make sure they put a package in
    If lRet = NOERRORS Then
        If Not oPackages Is Nothing Then
            'connect to the server that they supplied
            lRet = Connect(NodeValue("//ServerName"))
        Else
            Err.Raise -1, PROC_NAME, "No packages supplied."
        End If
    End If
    
    'get the packages collection
    If lRet = NOERRORS Then
        Set oMTSPackages = m_oCatalog.GetCollection("Packages")
    End If
    
    'populate the package
    If lRet = NOERRORS Then
        oMTSPackages.Populate
    End If
    
Exit_Point:
    'return the results
    GetPackages = lRet
    Set r_oMTSPackages = oMTSPackages
    Set r_oPackages = oPackages
    
    'clean up
    Set oMTSPackages = Nothing
    Set oPackages = Nothing
    
    'remove us from the error stack
    m_oErr.RemoveProcedureName PROC_NAME
    Exit Function
ErrorHandler:
    'handle the error
    lRet = Err.Number
    
    m_oErr.HandleError lRet, Err.Description
    GoTo Exit_Point
End Function

Private Function FindPackage(ByVal p_oMTSPackages As MTSAdmin.CatalogCollection, ByVal p_oPackage As IXMLDOMNode, ByRef r_lPackageID As Long, Optional ByRef r_oPackage As MTSAdmin.CatalogObject) As Long
    Dim lRet As Long
    Dim i As Long
    Dim oPackage As MTSAdmin.CatalogObject
    Dim lPackageID As Long
    Const PROC_NAME As String = "FindPackage"
    On Error GoTo ErrorHandler
    
    'add us to the error stack
    m_oErr.AddProcedureName PROC_NAME
        
    'loop through the packages
    For i = 0 To p_oMTSPackages.Count - 1
        'if they passed an id, use it (that's the only way to change the package name)
        If NodeValue("ID", p_oPackage) = "" Then
            'if the names match, use them
            If p_oMTSPackages.Item(i).Name = NodeValue("Name", p_oPackage) Then
                Set oPackage = p_oMTSPackages.Item(i)
                lPackageID = i
                Exit For
            End If
        Else
            If p_oMTSPackages.Item(i).Value("ID") = NodeValue("ID", p_oPackage) Then
                Set oPackage = p_oMTSPackages.Item(i)
                lPackageID = i
                Exit For
            End If
        End If
    Next i

    'check to see if we found the package
    If oPackage Is Nothing Then
        Err.Raise -2, PROC_NAME, "Package not found.  The name supplied was " & NodeValue("Name", p_oPackage) & ".  The ID supplied was " & NodeValue("ID", p_oPackage) & "."
    End If

Exit_Point:
    'return the results
    FindPackage = lRet
    Set r_oPackage = oPackage
    r_lPackageID = lPackageID
    
    'clean up
    Set oPackage = Nothing
    
    'remove us from the error stack
    m_oErr.RemoveProcedureName PROC_NAME
    Exit Function
ErrorHandler:
    'handle the error
    lRet = Err.Number
    
    m_oErr.HandleError lRet, Err.Description
    GoTo Exit_Point
End Function
Private Function FindComponent(ByVal p_oMTSComponents As MTSAdmin.CatalogCollection, ByVal p_oComponent As IXMLDOMNode, ByRef r_lComponentID As Long, Optional ByRef r_oComponent As MTSAdmin.CatalogObject) As Long
    Dim lRet As Long
    Dim i As Long
    Dim oComponent As MTSAdmin.CatalogObject
    Dim lComponentID As Long
    Const PROC_NAME As String = "FindComponent"
    On Error GoTo ErrorHandler
    
    'add us to the error stack
    m_oErr.AddProcedureName PROC_NAME
        
    'loop through the Components
    For i = 0 To p_oMTSComponents.Count - 1
        'if they passed an id, use it (that's the only way to change the Component name)
        If NodeValue("CLSID", p_oComponent) = "" Then
            'if the names match, use them
            If p_oMTSComponents.Item(i).Name = NodeValue("Name", p_oComponent) Then
                Set oComponent = p_oMTSComponents.Item(i)
                lComponentID = i
                Exit For
            End If
        Else
            If p_oMTSComponents.Item(i).Value("CLSID") = NodeValue("CLSID", p_oComponent) Then
                Set oComponent = p_oMTSComponents.Item(i)
                lComponentID = i
                Exit For
            End If
        End If
    Next i

    'check to see if we found the Component
    If oComponent Is Nothing Then
        Err.Raise -4, PROC_NAME, "Component not found.  The name supplied was " & NodeValue("Name", p_oComponent) & ".  The CLSID supplied was " & NodeValue("CLSID", p_oComponent) & "."
    End If

Exit_Point:
    'return the results
    FindComponent = lRet
    Set r_oComponent = oComponent
    r_lComponentID = lComponentID
    
    'clean up
    Set oComponent = Nothing
    
    'remove us from the error stack
    m_oErr.RemoveProcedureName PROC_NAME
    Exit Function
ErrorHandler:
    'handle the error
    lRet = Err.Number
    
    m_oErr.HandleError lRet, Err.Description
    GoTo Exit_Point
End Function

Private Function GetComponents(ByRef r_oMTSComponents As MTSAdmin.CatalogCollection, ByRef r_oComponents As IXMLDOMNodeList, ByVal p_oPackage As IXMLDOMNode, ByVal p_oMTSPackages As MTSAdmin.CatalogCollection, ByVal p_lPackageID As String) As Long
    Dim lRet As Long
    Dim oComponents As IXMLDOMNodeList
    Dim oMTSComponents As MTSAdmin.CatalogCollection
    Const PROC_NAME As String = "GetComponents"
    On Error GoTo ErrorHandler
    
    'add us to the error stack
    m_oErr.AddProcedureName PROC_NAME
    
    'get all of the Components they want to add--we only allow them to do one server at a time.
    If lRet = NOERRORS Then
        Set oComponents = p_oPackage.selectNodes("//Component")
    End If
    
    'make sure they put a Component in
    If lRet = NOERRORS Then
        If oComponents Is Nothing Then
            Err.Raise -3, PROC_NAME, "No Components supplied."
        End If
    End If
    
    'get the Components collection
    If lRet = NOERRORS Then
        Set oMTSComponents = p_oMTSPackages.GetCollection("ComponentsInPackage", p_lPackageID)
    End If
    
    'populate the Component
    If lRet = NOERRORS Then
        oMTSComponents.Populate
    End If
    
Exit_Point:
    'return the results
    GetComponents = lRet
    Set r_oMTSComponents = oMTSComponents
    Set r_oComponents = oComponents
    
    'clean up
    Set oMTSComponents = Nothing
    Set oComponents = Nothing
    
    'remove us from the error stack
    m_oErr.RemoveProcedureName PROC_NAME
    Exit Function
ErrorHandler:
    'handle the error
    lRet = Err.Number
    
    m_oErr.HandleError lRet, Err.Description
    GoTo Exit_Point
End Function

Private Function AddXML(ByRef r_oParent As IXMLDOMNode, ByVal p_sName As String, ByVal p_oPropertyCollection As MTSAdmin.CatalogCollection, ByVal p_oObject As MTSAdmin.CatalogObject) As Long
    Dim lRet As Long
    Dim oParent As IXMLDOMNode
    Dim i As Long
    Const PROC_NAME As String = "AddXML"
    On Error GoTo ErrorHandler
    
    'add us to the error stack
    m_oErr.AddProcedureName PROC_NAME
    
    'add the componentnode
    If lRet = NOERRORS Then
        lRet = AddElement(r_oParent, p_sName, "", oParent)
    End If
    
    If lRet = NOERRORS Then
        'make sure the property colleciton is populated
        p_oPropertyCollection.Populate
    End If
    'add the properties
    If lRet = NOERRORS Then
        'at a minimum get the name
        If p_oPropertyCollection.Count = 0 Then
            lRet = AddElement(oParent, "Name", p_oObject.Value("Name"))
        End If
        For i = 0 To p_oPropertyCollection.Count - 1
            If lRet = NOERRORS Then
                'make sure we don't have a write only value
                If Not p_oObject.IsPropertyWriteOnly(p_oPropertyCollection.Item(i).Name) Then
                    lRet = AddElement(oParent, p_oPropertyCollection.Item(i).Name, p_oObject.Value(p_oPropertyCollection.Item(i).Name))
                End If
            End If
        Next i
    End If
    
Exit_Point:
    'return the results
    AddXML = lRet
    
    'remove us from the error stack
    m_oErr.RemoveProcedureName PROC_NAME
    Exit Function
ErrorHandler:
    'handle the error
    lRet = Err.Number
    
    m_oErr.HandleError lRet, Err.Description
    GoTo Exit_Point
End Function
Private Function AddElement(ByRef r_oParent As IXMLDOMNode, ByVal p_sNodeName As String, ByVal p_sNodeValue As String, Optional ByRef r_oNode As IXMLDOMNode) As Long
    Dim lRet As Long
    Dim oNode As IXMLDOMNode
    Const PROC_NAME As String = "AddElement"
    On Error GoTo ErrorHandler
    
    'add us to the error stack
    m_oErr.AddProcedureName PROC_NAME
    
    'create the node
    Set oNode = m_oDOM.createElement(p_sNodeName)
    
    'set it's text
    oNode.Text = p_sNodeValue
    
    'append it to the parent
    r_oParent.appendChild oNode
    
Exit_Point:
    'return the results
    AddElement = lRet
    Set r_oNode = oNode
    
    'clean up
    Set oNode = Nothing
    
    'remove us from the error stack
    m_oErr.RemoveProcedureName PROC_NAME
    Exit Function
ErrorHandler:
    'handle the error
    lRet = Err.Number
    
    m_oErr.HandleError lRet, Err.Description
    GoTo Exit_Point
End Function


Private Function SetValues(ByRef r_oCollection As MTSAdmin.CatalogCollection, ByVal lID As Long, p_oNode As IXMLDOMNode) As Long
    Dim lRet As Long
    Dim oPropertyInfo As MTSAdmin.CatalogCollection
    Dim i As Long
    Const PROC_NAME As String = "SetValues"
    On Error GoTo ErrorHandler
    
    'add us to the error stack
    m_oErr.AddProcedureName PROC_NAME
    
    'get the property info
    Set oPropertyInfo = r_oCollection.GetCollection("PropertyInfo", r_oCollection.Item(lID).Key)
    
    oPropertyInfo.Populate
    
    'loop through the props and set them
    For i = 0 To oPropertyInfo.Count - 1
        If NodeExists(oPropertyInfo.Item(i).Name, p_oNode) Then
            SetValue r_oCollection.Item(lID), p_oNode, oPropertyInfo.Item(i).Name
        End If
    Next i
    
    
Exit_Point:
    'return the results
    SetValues = lRet
    
    'remove us from the error stack
    m_oErr.RemoveProcedureName PROC_NAME
    Exit Function
ErrorHandler:
    'handle the error
    lRet = Err.Number
    
    m_oErr.HandleError lRet, Err.Description
    GoTo Exit_Point
End Function

