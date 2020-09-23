VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmMTSAdmin 
   Caption         =   "MTS Admin"
   ClientHeight    =   6750
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10305
   Icon            =   "frmMTSAdmin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6750
   ScaleWidth      =   10305
   Begin MSComctlLib.ImageList ImageList32 
      Left            =   4980
      Top             =   5940
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMTSAdmin.frx":030A
            Key             =   "Component"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMTSAdmin.frx":096C
            Key             =   "Computer"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMTSAdmin.frx":0E1C
            Key             =   "Folder"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMTSAdmin.frx":1216
            Key             =   "Interface"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMTSAdmin.frx":17A1
            Key             =   "Method"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMTSAdmin.frx":1D6C
            Key             =   "Package"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMTSAdmin.frx":2274
            Key             =   "RemoteComputer"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList16 
      Left            =   3900
      Top             =   5940
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMTSAdmin.frx":26A9
            Key             =   "Component"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMTSAdmin.frx":2AD3
            Key             =   "Services"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMTSAdmin.frx":2E9A
            Key             =   "Computer"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMTSAdmin.frx":3268
            Key             =   "Folder"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMTSAdmin.frx":35FE
            Key             =   "Interface"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMTSAdmin.frx":3A02
            Key             =   "Method"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMTSAdmin.frx":3DF8
            Key             =   "Package"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMTSAdmin.frx":41B6
            Key             =   "RemoteComputer"
         EndProperty
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid grdDetails 
      Height          =   6195
      Left            =   3840
      TabIndex        =   1
      Top             =   420
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   10927
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      BackColorFixed  =   -2147483644
      BackColorSel    =   -2147483643
      BackColorBkg    =   16777215
      AllowBigSelection=   0   'False
      HighLight       =   2
      GridLines       =   0
      MergeCells      =   1
      AllowUserResizing=   3
   End
   Begin MSComctlLib.TreeView treMTS 
      Height          =   6255
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   11033
      _Version        =   393217
      Indentation     =   530
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      ImageList       =   "ImageList16"
      Appearance      =   1
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuNew 
      Caption         =   "New Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuPackageNew 
         Caption         =   "New Package . . ."
      End
   End
   Begin VB.Menu mnuPackageMenu 
      Caption         =   "Package Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuComponentNew 
         Caption         =   "New Component . . ."
      End
      Begin VB.Menu sep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExport 
         Caption         =   "Export . . ."
      End
      Begin VB.Menu mnuShutDown 
         Caption         =   "Shut Down"
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPackageDelete 
         Caption         =   "Delete"
      End
      Begin VB.Menu sep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPackageProperties 
         Caption         =   "Properties"
      End
   End
   Begin VB.Menu mnuComponentMenu 
      Caption         =   "Component Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuComponentDelete 
         Caption         =   "Delete"
      End
      Begin VB.Menu sep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuComponentProperties 
         Caption         =   "Properties"
      End
   End
   Begin VB.Menu mnuComputers 
      Caption         =   "Computer Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuAddComputer 
         Caption         =   "Add Computer"
      End
   End
   Begin VB.Menu mnuComputer2 
      Caption         =   "Computer Menu2"
      Visible         =   0   'False
      Begin VB.Menu mnuComputerDelete 
         Caption         =   "Delete"
      End
   End
End
Attribute VB_Name = "frmMTSAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_oRoot As MSComctlLib.Node
Private m_oCurrentNode As MSComctlLib.Node
Private m_oDOM As DOMDocument30

Private m_bProcessExpand As Boolean
Private m_bMouseDown As Boolean
Private m_lLeft As Long
Private m_oErr As CError
Public Property Get CurrentNode() As MSComctlLib.Node
    Set CurrentNode = m_oCurrentNode
End Property

Public Sub RefreshTree(Optional ByRef r_oNode As MSComctlLib.Node)
    Dim oParent As MSComctlLib.Node
    'get the node to refresh
    
    If r_oNode Is Nothing Then
        Set oParent = m_oCurrentNode
    Else
        Set oParent = r_oNode
    End If

    If Not oParent Is Nothing Then
        If oParent.Key = "Root" Then
            oParent.Sorted = True
        Else
            m_bProcessExpand = False
            While oParent.children > 0
                treMTS.Nodes.Remove oParent.Child.Index
            Wend
            
            'add a blank child
            AddBlankNode oParent
            
            'force the expand
            m_bProcessExpand = True
            oParent.Expanded = True
        End If
    End If
    Set oParent = Nothing
End Sub
Private Sub Form_Load()
    Dim lRet As Long
    Const PROC_NAME As String = "Form_Load"
    Set m_oErr = New CError
    
    On Error GoTo ErrorHandler
    
    m_oErr.AddProcedureName PROC_NAME
    
    m_bProcessExpand = False
    'build the start nodes
    Set m_oRoot = treMTS.Nodes.Add(, , "Root", "Computers")
    m_oRoot.Image = "Folder"
    AddBlankNode m_oRoot
    
    'make sure we aren't showing any nodes
    m_oRoot.Expanded = False
    m_bProcessExpand = True
    
    InitDOM m_oDOM
    m_lLeft = 4000
    
    GetSettings
Exit_Point:
    'return the results
    
    'clean up
    
    'remove us from the error stack
    m_oErr.RemoveProcedureName PROC_NAME
    Exit Sub
ErrorHandler:
    'handle the error
    lRet = Err.Number
    
    m_oErr.HandleError lRet, Err.Description, p_bDisplayMessage:=True
    GoTo Exit_Point
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If x >= m_lLeft And x <= m_lLeft + 50 Then
        Me.MousePointer = vbSizeWE
    Else
        Me.MousePointer = vbDefault
    End If
    If Button = 1 Then
        m_lLeft = x
        Form_Resize
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SaveSettings
    Set m_oRoot = Nothing
    Set m_oCurrentNode = Nothing
    Set m_oDOM = Nothing
    Set m_oErr = Nothing
End Sub

Private Sub Form_Resize()
    If Not frmMTSAdmin.WindowState = vbMinimized Then
        If m_lLeft < 100 Then
            m_lLeft = 100
        End If
        If m_lLeft > frmMTSAdmin.ScaleWidth - 100 Then
            m_lLeft = frmMTSAdmin.ScaleWidth - 100
        End If
        
        'fraSpacer.Height = frmMTSAdmin.ScaleHeight
        
        'move the trecontrol
        treMTS.Move 0, 0, m_lLeft, frmMTSAdmin.ScaleHeight
        
        'move the grid
        grdDetails.Move m_lLeft + 50, 0, frmMTSAdmin.ScaleWidth - m_lLeft - 50, frmMTSAdmin.ScaleHeight
        
    End If
End Sub



Private Sub grdDetails_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Me.MousePointer = vbSizeWE Then
        Me.MousePointer = vbDefault
    End If

End Sub

Private Sub mnuAddComputer_Click()
    Const PROC_NAME As String = "mnuAddComputer_Click"
    Dim lRet As Long
    Dim oNode As MSComctlLib.Node
    
    On Error GoTo ErrorHandler
    
    m_oErr.AddProcedureName PROC_NAME
    
    AddComputer ""
    
    
Exit_Point:
    'return the results
    
    'clean up
    Set oNode = Nothing
    m_bProcessExpand = True
    
    'remove us from the error stack
    m_oErr.RemoveProcedureName PROC_NAME
    Exit Sub
ErrorHandler:
    'handle the error
    lRet = Err.Number
    
    m_oErr.HandleError lRet, Err.Description, p_bDisplayMessage:=True
    GoTo Exit_Point
End Sub

Private Sub mnuComponentDelete_Click()
    Dim lPos As Long
    Dim sCLSID As String
    Dim oNode As MSComctlLib.Node
    Dim sKey As String
    Dim lRet As Long
    Const PROC_NAME As String = "mnuComponentDelete_Click"
    On Error GoTo ErrorHandler
    m_oErr.AddProcedureName PROC_NAME
    
    lRet = MsgBox("Are you sure you want to delete " & m_oCurrentNode.Text & "?", vbCritical + vbYesNo, "Delete Component")
    If lRet = vbYes Then
        Me.MousePointer = vbHourglass
        lPos = InStrRev(m_oCurrentNode.Key, ":")
        sCLSID = Mid$(m_oCurrentNode.Key, lPos + 1)
        If m_oCurrentNode.Parent.Parent.Parent.Text = "My Computer" Then
            DeleteComponent "", m_oCurrentNode.Parent.Text, sCLSID
        Else
            DeleteComponent m_oCurrentNode.Parent.Parent.Parent.Text, m_oCurrentNode.Parent.Text, sCLSID
        End If
        
        RefreshTree m_oCurrentNode.Parent
    End If
    
Exit_Point:
    'return the results
    
    'clean up
    
    'remove us from the error stack
    m_oErr.RemoveProcedureName PROC_NAME
    Exit Sub
ErrorHandler:
    'handle the error
    lRet = Err.Number
    
    m_oErr.HandleError lRet, Err.Description, p_bDisplayMessage:=True
    GoTo Exit_Point
End Sub

Private Sub mnuComponentNew_Click()
    Dim oTempDOM As DOMDocument30
    Dim oPackage As IXMLDOMNode
    Dim oNode As IXMLDOMNode
    Dim lRet As Long
    Const PROC_NAME As String = "mnuComponentNew_Click"
    On Error GoTo ErrorHandler
    m_oErr.AddProcedureName PROC_NAME
    
    'init the blank dom
    InitDOM oTempDOM
    
    'load the dom with blank xml
    m_oDOM.loadXML XMLSTRING
    
    'load the current node xml
    oTempDOM.loadXML m_oCurrentNode.Tag
    
    'get the package node
    Set oPackage = m_oDOM.selectSingleNode("/Root/Package")
    
    'set the new node
    Set oNode = oTempDOM.selectSingleNode("/Package")
    
    'replace the xml
    m_oDOM.selectSingleNode("/Root").replaceChild oNode, oPackage
            
    'set the servername
    If m_oCurrentNode.Parent.Parent.Text = "My Computer" Then
        SetNodeValue "//ServerName", ""
    Else
        SetNodeValue "//ServerName", m_oCurrentNode.Parent.Parent.Text
    End If
    
    'load the form
    Load frmAddComponent
    
    'set the dom
    Set frmAddComponent.DOM = m_oDOM
    
    'show the form
    frmAddComponent.Show
    
Exit_Point:
    'return the results
    
    'clean up
    Set oNode = Nothing
    Set oPackage = Nothing
    Set oTempDOM = Nothing
    
    'remove us from the error stack
    m_oErr.RemoveProcedureName PROC_NAME
    Exit Sub
ErrorHandler:
    'handle the error
    lRet = Err.Number
    
    m_oErr.HandleError lRet, Err.Description, p_bDisplayMessage:=True
    GoTo Exit_Point
End Sub

Private Sub mnuComponentProperties_Click()
    Dim lRet As Long
    Const PROC_NAME As String = "mnuComponentProperties"
    On Error GoTo ErrorHandler
    
    m_oErr.AddProcedureName PROC_NAME
    
    'get the xml from the tag
    m_oDOM.loadXML m_oCurrentNode.Tag
    
    'load the properties form
    Load frmComponentProperties
    
    'get the server name and set it
    If m_oCurrentNode.Parent.Parent.Parent.Text = "My Computer" Then
        frmComponentProperties.ServerName = ""
    Else
        frmComponentProperties.ServerName = m_oCurrentNode.Parent.Parent.Parent.Text
    End If
    
    'set the dom
    Set frmComponentProperties.DOM = m_oDOM
    
    'show the form
    frmComponentProperties.Show
Exit_Point:
    'return the results
    
    'clean up
    
    'remove us from the error stack
    m_oErr.RemoveProcedureName PROC_NAME
    Exit Sub
ErrorHandler:
    'handle the error
    lRet = Err.Number
    
    m_oErr.HandleError lRet, Err.Description, p_bDisplayMessage:=True
    GoTo Exit_Point
End Sub

Private Sub mnuComputerDelete_Click()
    Dim lRet As Long
    Dim oReg As CRegistry
    If m_oCurrentNode.Text <> "My Computer" Then
        Set oReg = New CRegistry
        lRet = oReg.OpenRegistry(HKEY_LOCAL_MACHINE, "SOFTWARE\WOW\Client\Computers")
        If lRet Then
            oReg.DeleteValue m_oCurrentNode.Text
            treMTS.Nodes.Remove m_oCurrentNode.Index
        End If
        oReg.CloseRegistry
        Set oReg = Nothing
    End If
End Sub

Private Sub mnuExit_Click()
    Unload frmMTSAdmin
End Sub

Private Sub mnuExport_Click()
    Dim oTempDOM As DOMDocument30
    Dim oPackage As IXMLDOMNode
    Dim oNode As IXMLDOMNode
    Dim lRet As Long
    Const PROC_NAME As String = "mnuExport_Click"
    On Error GoTo ErrorHandler
    m_oErr.AddProcedureName PROC_NAME
    'init the blank dom
    InitDOM oTempDOM
    
    'load the dom with blank xml
    m_oDOM.loadXML XMLSTRING
    
    'load the current node xml
    oTempDOM.loadXML m_oCurrentNode.Tag
    
    'get the package node
    Set oPackage = m_oDOM.selectSingleNode("/Root/Package")
    
    'set the new node
    Set oNode = oTempDOM.selectSingleNode("/Package")
    
    'replace the xml
    m_oDOM.selectSingleNode("/Root").replaceChild oNode, oPackage
            
    'set the servername
    If m_oCurrentNode.Parent.Parent.Text = "My Computer" Then
        SetNodeValue "//ServerName", ""
    Else
        SetNodeValue "//ServerName", m_oCurrentNode.Parent.Parent.Text
    End If
    
    'load the form
    Load frmExport
    
    'set the dom
    Set frmExport.DOM = m_oDOM
    
    'show the form
    frmExport.Show
    
Exit_Point:
    'return the results
    
    'clean up
    Set oNode = Nothing
    Set oPackage = Nothing
    Set oTempDOM = Nothing
    
    'remove us from the error stack
    m_oErr.RemoveProcedureName PROC_NAME
    Exit Sub
ErrorHandler:
    'handle the error
    lRet = Err.Number
    
    m_oErr.HandleError lRet, Err.Description, p_bDisplayMessage:=True
    GoTo Exit_Point
End Sub

Private Sub mnuPackageDelete_Click()
    Dim lPos As Long
    Dim sCLSID As String
    Dim oNode As MSComctlLib.Node
    Dim sKey As String
    Dim lRet As Long
    Const PROC_NAME As String = "mnuPackageDelete_Click"
    On Error GoTo ErrorHandler
    m_oErr.AddProcedureName PROC_NAME
    
    'make sure they really want to do this
    lRet = MsgBox("Are you sure you want to delete " & m_oCurrentNode.Text & "?", vbCritical + vbYesNo, "Delete Package")
    
    'if they are sure, then run the delete code
    If lRet = vbYes Then
        Me.MousePointer = vbHourglass
        lPos = InStrRev(m_oCurrentNode.Key, ":")
        sCLSID = Mid$(m_oCurrentNode.Key, lPos + 1)
        If m_oCurrentNode.Parent.Parent.Text = "My Computer" Then
            DeletePackage "", sCLSID
        Else
            DeletePackage m_oCurrentNode.Parent.Parent.Text, sCLSID
        End If
        
        RefreshTree m_oCurrentNode.Parent
    End If
Exit_Point:
    'return the results
    
    'clean up
    
    'remove us from the error stack
    m_oErr.RemoveProcedureName PROC_NAME
    Exit Sub
ErrorHandler:
    'handle the error
    lRet = Err.Number
    
    m_oErr.HandleError lRet, Err.Description, p_bDisplayMessage:=True
    GoTo Exit_Point
End Sub

Private Sub mnuPackageNew_Click()
    Dim lRet As Long
    Const PROC_NAME As String = "mnuPackageNew_Click"
    On Error GoTo ErrorHandler
    m_oErr.AddProcedureName PROC_NAME
    
    'load the dom with blank xml
    m_oDOM.loadXML XMLSTRING
                
    'set the servername
    If m_oCurrentNode.Parent.Text = "My Computer" Then
        SetNodeValue "//ServerName", ""
    Else
        SetNodeValue "//ServerName", m_oCurrentNode.Parent.Text
    End If
    
    'load the form
    Load frmAddPackage
    
    'set the dom
    Set frmAddPackage.DOM = m_oDOM
    
    'show the form
    frmAddPackage.Show
Exit_Point:
    'return the results
    
    'clean up
    
    'remove us from the error stack
    m_oErr.RemoveProcedureName PROC_NAME
    Exit Sub
ErrorHandler:
    'handle the error
    lRet = Err.Number
    
    m_oErr.HandleError lRet, Err.Description, p_bDisplayMessage:=True
    GoTo Exit_Point
End Sub

Private Sub mnuPackageProperties_Click()
    Dim lRet As Long
    
    Const PROC_NAME As String = "mnuPackageProperties_Click"
    On Error GoTo ErrorHandler
    m_oErr.AddProcedureName PROC_NAME


    m_oDOM.loadXML m_oCurrentNode.Tag
    
    Load frmPackageProperties
    
    If m_oCurrentNode.Parent.Parent.Text = "My Computer" Then
        frmPackageProperties.ServerName = ""
    Else
        frmPackageProperties.ServerName = m_oCurrentNode.Parent.Parent.Text
    End If
    
    Set frmPackageProperties.DOM = m_oDOM
    frmPackageProperties.Show
Exit_Point:
    'return the results
    
    'clean up
    
    'remove us from the error stack
    m_oErr.RemoveProcedureName PROC_NAME
    Exit Sub
ErrorHandler:
    'handle the error
    lRet = Err.Number
    
    m_oErr.HandleError lRet, Err.Description, p_bDisplayMessage:=True
    GoTo Exit_Point
    

End Sub

Private Sub mnuShutDown_Click()
    Dim lRet As Long

    Const PROC_NAME As String = "mnuShutDown_Click"
    On Error GoTo ErrorHandler
    m_oErr.AddProcedureName PROC_NAME

    Me.MousePointer = vbHourglass
    If m_oCurrentNode.Parent.Parent.Text = "My Computer" Then
        ShutdownPackage "", m_oCurrentNode.Text
    Else
        ShutdownPackage m_oCurrentNode.Parent.Parent.Text, m_oCurrentNode.Text
    End If
    Me.MousePointer = vbDefault
Exit_Point:
    'return the results
    
    'clean up
    
    'remove us from the error stack
    m_oErr.RemoveProcedureName PROC_NAME
    Exit Sub
ErrorHandler:
    'handle the error
    lRet = Err.Number
    
    m_oErr.HandleError lRet, Err.Description, p_bDisplayMessage:=True
    GoTo Exit_Point
    
End Sub

Private Sub treMTS_BeforeLabelEdit(Cancel As Integer)
    If treMTS.SelectedItem.Text = "My Computer" Then
        Cancel = True
    End If
    
End Sub

Private Sub treMTS_Expand(ByVal Node As MSComctlLib.Node)
    Dim oNode As MSComctlLib.Node
    Dim lRet As Long
    Dim oList As IXMLDOMNodeList
    Dim i As Long
    Const PROC_NAME As String = "treMTS_Expand"
    On Error GoTo ErrorHandler
    m_oErr.AddProcedureName PROC_NAME
    
    If m_bProcessExpand Then
        Me.MousePointer = vbHourglass
        LockWindowUpdate treMTS.hWnd
        m_bProcessExpand = False
        Select Case Left(Node.Key, 8)
            Case "Root"
                If Not Left(Node.Child.Key, 8) = "Computer" Then
                    'kill the blank node
                    treMTS.Nodes.Remove Node.Child.Key
                    
                    'create the new node
                    Set oNode = treMTS.Nodes.Add
                    
                    'set it's parent
                    Set oNode.Parent = m_oRoot
                    
                    'set its key
                    oNode.Key = "Computer1"
                    
                    'set it's text
                    oNode.Text = "My Computer"
                    
                    'set it's icon
                    oNode.Image = "Computer"
                    
                    AddBlankNode oNode
                End If
            Case "Computer"
                If Not Left(Node.Child.Key, 8) = "PackFold" Then
                    'kill the blank node
                    treMTS.Nodes.Remove Node.Child.Key
                    
                    'create the new node
                    Set oNode = treMTS.Nodes.Add
                    
                    'set it's parent
                    Set oNode.Parent = Node
                    
                    'set its key
                    oNode.Key = "PackFold:" & Node.Key
                    
                    'set it's text
                    oNode.Text = "Packages Installed"
                    
                    'set it's icon
                    oNode.Image = "Folder"
                    
                    'change the blank node's location
                    AddBlankNode oNode
                End If
            Case "PackFold"
                If Not Left(Node.Child.Key, 8) = "Package" Then
                    'kill the blank node
                    treMTS.Nodes.Remove Node.Child.Key
                    
                    'get the packages for this computer
                    If Node.Parent.Text = "My Computer" Then
                        lRet = GetPackages("", Node, oList)
                    Else
                        lRet = GetPackages(Node.Parent.Text, Node, oList)
                    End If
                    For i = 0 To oList.length - 1
                        Set oNode = treMTS.Nodes.Add
                        oNode.Key = "Package:" & Node.Parent.Key & ":" & oList.Item(i).selectSingleNode("ID").Text
                        oNode.Text = oList.Item(i).selectSingleNode("Name").Text
                        Set oNode.Parent = Node
                        oNode.Tag = oList.Item(i).xml
                        'set it's icon
                        oNode.Image = "Package"
                        AddBlankNode oNode
                    Next i
    
                    'sort the tree
                    Node.Sorted = True
                    
                End If
            Case "Package:"
                If Not Left(Node.Child.Key, 8) = "Computer" Then
                    'kill the blank node
                    treMTS.Nodes.Remove Node.Child.Key
                    
                    'get the packages for this computer
                    'get the packages for this computer
                    If Node.Parent.Parent.Text = "My Computer" Then
                        lRet = GetComponents("", Node, oList)
                    Else
                        lRet = GetComponents(Node.Parent.Parent.Text, Node, oList)
                    End If
                    
                    For i = 0 To oList.length - 1
                        Set oNode = treMTS.Nodes.Add
                        oNode.Key = "Component:" & Node.Parent.Parent.Key & ":" & oList.Item(i).selectSingleNode("CLSID").Text
                        oNode.Text = oList.Item(i).selectSingleNode("ProgID").Text
                        oNode.Tag = oList.Item(i).xml
                        Set oNode.Parent = Node
                        'set it's icon
                        oNode.Image = "Component"
                        'AddBlankNode oNode
                    Next i
    
                    'sort the tree
                    Node.Sorted = True
                    
                End If
            Case Else
        End Select
    End If
Exit_Point:
    'return the results
    
    'clean up
    Set oNode = Nothing
    Set oList = Nothing
    Me.MousePointer = vbDefault
    m_bProcessExpand = True
    LockWindowUpdate 0
    
    'remove us from the error stack
    m_oErr.RemoveProcedureName PROC_NAME
    Exit Sub
ErrorHandler:
    'handle the error
    lRet = Err.Number
    
    m_oErr.HandleError lRet, Err.Description, p_bDisplayMessage:=True
    GoTo Exit_Point
End Sub

Private Sub treMTS_KeyDown(KeyCode As Integer, Shift As Integer)
    Const F5 As Long = 116
    If KeyCode = F5 Then
        RefreshTree
    End If
End Sub

Private Sub treMTS_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    'find out what node they clicked on
    Set m_oCurrentNode = treMTS.HitTest(x, y)
    If Not m_oCurrentNode Is Nothing And Button = vbRightButton Then
        m_oCurrentNode.Selected = True
        Select Case Left(m_oCurrentNode.Key, 8)
            Case "Root"
                PopupMenu mnuComputers
            Case "Package:"
                PopupMenu mnuPackageMenu
            Case "Componen"
                PopupMenu mnuComponentMenu
            Case "PackFold"
                PopupMenu mnuNew
            Case "Computer"
                PopupMenu mnuComputer2
        End Select
            
    End If
    
End Sub
Private Function AddBlankNode(ByRef r_oParent As MSComctlLib.Node) As Long
    Dim lRet As Long
    Dim oNode As MSComctlLib.Node
    Const PROC_NAME As String = "AddBlankNode"
    On Error GoTo ErrorHandler
    m_oErr.AddProcedureName PROC_NAME

    m_bProcessExpand = False
    Set oNode = treMTS.Nodes.Add
    
    oNode.Text = ""
    oNode.Key = "Blank:" & r_oParent.Key
    
    Set oNode.Parent = r_oParent
    
    r_oParent.Expanded = False
Exit_Point:
    'return the results
    AddBlankNode = lRet
    
    'clean up
    Set oNode = Nothing
    m_bProcessExpand = True
    
    'remove us from the error stack
    m_oErr.RemoveProcedureName PROC_NAME
    Exit Function
ErrorHandler:
    'handle the error
    lRet = Err.Number
    
    m_oErr.HandleError lRet, Err.Description, p_bDisplayMessage:=True
    GoTo Exit_Point
    
End Function
Private Function GetPackages(ByVal p_sComputer As String, ByRef r_oParentNode As MSComctlLib.Node, ByRef r_oList As IXMLDOMNodeList) As Long
    Dim lRet As Long
    Dim sXML As String
    Dim oMTSAdmin As WOWMTSAdmin.CMtsAdmin
    Dim oPackages As IXMLDOMNodeList
    Dim oNode As MSComctlLib.Node
    Dim oXMLNode As IXMLDOMNode
    Dim i As Long
    
    Const PROC_NAME As String = "GetPackages"
    On Error GoTo ErrorHandler
    m_oErr.AddProcedureName PROC_NAME
    
    'load the dom
    m_oDOM.loadXML XMLSTRING
    
    'set the computername
    SetNodeValue "/Root/ServerName", p_sComputer
    
    'remove the package node
    Set oXMLNode = m_oDOM.selectSingleNode("/Root/Package")
    m_oDOM.selectSingleNode("/Root").removeChild oXMLNode
    
    'create an instance of the mts admin on the computer supplied
    If p_sComputer = "" Then
        Set oMTSAdmin = CreateObject("WOWMTSAdmin.CMtsAdmin")
    Else
        Set oMTSAdmin = CreateObject("WOWMTSAdmin.CMtsAdmin", p_sComputer)
    End If
    
    'get the packages
    lRet = oMTSAdmin.ListPackages(m_oDOM.xml, sXML)
    
    Set oMTSAdmin = Nothing
    
    'if we didn't have an error, load the xml
    If lRet = NOERRORS Then
        m_oDOM.loadXML sXML
    End If
    
    'get the package list
    Set oPackages = m_oDOM.selectNodes("//Package")
        
Exit_Point:
    'return the results
    GetPackages = lRet
    Set r_oList = oPackages
    
    'clean up
    Set oMTSAdmin = Nothing
    Set oPackages = Nothing
    Set oNode = Nothing
    Set oXMLNode = Nothing
    
    'remove us from the error stack
    m_oErr.RemoveProcedureName PROC_NAME
    Exit Function
ErrorHandler:
    'handle the error
    lRet = Err.Number
    
    m_oErr.HandleError lRet, Err.Description, p_bDisplayMessage:=True
    GoTo Exit_Point
    
End Function

Private Function SetNodeValue(ByVal p_sNodeName As String, ByVal p_sValue As String, Optional ByRef p_oParent As IXMLDOMNode) As Long
    Dim oNode As IXMLDOMNode
    Dim oParent As IXMLDOMNode
    
    Dim lRet As Long

    Const PROC_NAME As String = "Stub"
    On Error GoTo ErrorHandler
    m_oErr.AddProcedureName PROC_NAME
    
    'check for a parent
    If p_oParent Is Nothing Then
        Set oParent = m_oDOM
    Else
        Set oParent = p_oParent
    End If
    
    'get the node
    Set oNode = oParent.selectSingleNode(p_sNodeName)
    
    'set it's value
    If Not oNode Is Nothing Then
        oNode.Text = p_sValue
    End If
    
Exit_Point:
    'return the results
    SetNodeValue = lRet
    
    'clean up
    Set oNode = Nothing
    
    'remove us from the error stack
    m_oErr.RemoveProcedureName PROC_NAME
    Exit Function
ErrorHandler:
    'handle the error
    lRet = Err.Number
    
    m_oErr.HandleError lRet, Err.Description, p_bDisplayMessage:=True
    GoTo Exit_Point
End Function
Private Function GetComponents(ByVal p_sComputer As String, ByRef r_oParentNode As MSComctlLib.Node, ByRef r_oComponents As IXMLDOMNodeList) As Long
    Dim lRet As Long
    Dim sXML As String
    Dim oMTSAdmin As WOWMTSAdmin.CMtsAdmin
    Dim oComponents As IXMLDOMNodeList
    Dim oNode As MSComctlLib.Node
    Dim oXMLNode As IXMLDOMNode
    Dim i As Long
    Const PROC_NAME As String = "GetComponents"
    On Error GoTo ErrorHandler
    m_oErr.AddProcedureName PROC_NAME
    
    'load the dom
    m_oDOM.loadXML XMLSTRING
    
    'set the computername
    SetNodeValue "/Root/ServerName", p_sComputer
    SetNodeValue "/Root/Package/Name", r_oParentNode.Text
    
    'remove the package node
    Set oXMLNode = m_oDOM.selectSingleNode("/Root/Package/Component")
    m_oDOM.selectSingleNode("/Root/Package").removeChild oXMLNode
    
    'create an instance of the mts admin on the computer supplied
    If p_sComputer = "" Then
        Set oMTSAdmin = CreateObject("WOWMTSAdmin.CMtsAdmin")
    Else
        Set oMTSAdmin = CreateObject("WOWMTSAdmin.CMtsAdmin", p_sComputer)
    End If
    
    'get the packages
    lRet = oMTSAdmin.ListComponents(m_oDOM.xml, sXML)
    
    Set oMTSAdmin = Nothing
    
    'if we didn't have an error, load the xml
    If lRet = NOERRORS Then
        m_oDOM.loadXML sXML
    End If
    
    'get the package list
    Set oComponents = m_oDOM.selectNodes("//Component")
        
Exit_Point:
    'return the results
    GetComponents = lRet
    Set r_oComponents = oComponents
    
    'clean up
    Set oMTSAdmin = Nothing
    Set oComponents = Nothing
    Set oNode = Nothing
    Set oXMLNode = Nothing
    
    'remove us from the error stack
    m_oErr.RemoveProcedureName PROC_NAME
    Exit Function
ErrorHandler:
    'handle the error
    lRet = Err.Number
    
    m_oErr.HandleError lRet, Err.Description, p_bDisplayMessage:=True
    GoTo Exit_Point
End Function
Private Function ShutdownPackage(ByVal p_sComputer As String, ByVal p_sPackage As String) As Long
    Dim lRet As Long
    Dim sXML As String
    Dim oMTSAdmin As WOWMTSAdmin.CMtsAdmin
    Const PROC_NAME As String = "ShutdownPackage"
    On Error GoTo ErrorHandler
    m_oErr.AddProcedureName PROC_NAME
    
    'load the dom
    m_oDOM.loadXML XMLSTRING
    
    'set the computername
    SetNodeValue "/Root/ServerName", p_sComputer
    SetNodeValue "/Root/Package/Name", p_sPackage
        
    'create an instance of the mts admin on the computer supplied
    If p_sComputer = "" Then
        Set oMTSAdmin = CreateObject("WOWMTSAdmin.CMtsAdmin")
    Else
        Set oMTSAdmin = CreateObject("WOWMTSAdmin.CMtsAdmin", p_sComputer)
    End If
    
    'get the packages
    lRet = oMTSAdmin.ShutdownPackage(m_oDOM.xml, sXML)
            
Exit_Point:
    'return the results
    ShutdownPackage = lRet
    
    'clean up
    Set oMTSAdmin = Nothing
    
    'remove us from the error stack
    m_oErr.RemoveProcedureName PROC_NAME
    Exit Function
ErrorHandler:
    'handle the error
    lRet = Err.Number
    
    m_oErr.HandleError lRet, Err.Description, p_bDisplayMessage:=True
    GoTo Exit_Point
    
End Function
Private Function DeleteComponent(ByVal p_sComputer As String, ByVal p_sPackage As String, ByVal p_sComponent As String) As Long
    Dim lRet As Long
    Dim sXML As String
    Dim oMTSAdmin As WOWMTSAdmin.CMtsAdmin
    Const PROC_NAME As String = "DeleteComponent"
    On Error GoTo ErrorHandler
    m_oErr.AddProcedureName PROC_NAME
    
    'load the dom
    m_oDOM.loadXML XMLSTRING
    
    'set the computername
    SetNodeValue "/Root/ServerName", p_sComputer
    SetNodeValue "/Root/Package/Name", p_sPackage
    SetNodeValue "/Root/Package/Component/CLSID", p_sComponent
        
    'create an instance of the mts admin on the computer supplied
    If p_sComputer = "" Then
        Set oMTSAdmin = CreateObject("WOWMTSAdmin.CMtsAdmin")
    Else
        Set oMTSAdmin = CreateObject("WOWMTSAdmin.CMtsAdmin", p_sComputer)
    End If
    
    'get the packages
    lRet = oMTSAdmin.RemoveComponent(m_oDOM.xml, sXML)
                
Exit_Point:
    'return the results
    DeleteComponent = lRet
    
    'clean up
    Set oMTSAdmin = Nothing
    
    'remove us from the error stack
    m_oErr.RemoveProcedureName PROC_NAME
    Exit Function
ErrorHandler:
    'handle the error
    lRet = Err.Number
    
    m_oErr.HandleError lRet, Err.Description, p_bDisplayMessage:=True
    GoTo Exit_Point
    
End Function
Private Function DeletePackage(ByVal p_sComputer As String, ByVal p_sPackageID As String) As Long
    Dim lRet As Long
    Dim sXML As String
    Dim oMTSAdmin As WOWMTSAdmin.CMtsAdmin
    Const PROC_NAME As String = "DeletePackage"
    On Error GoTo ErrorHandler
    m_oErr.AddProcedureName PROC_NAME
    
    'load the dom
    m_oDOM.loadXML XMLSTRING
    
    'set the computername
    SetNodeValue "/Root/ServerName", p_sComputer
    SetNodeValue "/Root/Package/ID", p_sPackageID
        
    'create an instance of the mts admin on the computer supplied
    If p_sComputer = "" Then
        Set oMTSAdmin = CreateObject("WOWMTSAdmin.CMtsAdmin")
    Else
        Set oMTSAdmin = CreateObject("WOWMTSAdmin.CMtsAdmin", p_sComputer)
    End If
    
    'get the packages
    lRet = oMTSAdmin.RemovePackage(m_oDOM.xml, sXML)
                
Exit_Point:
    'return the results
    
    'clean up
    Set oMTSAdmin = Nothing
    
    'remove us from the error stack
    m_oErr.RemoveProcedureName PROC_NAME
    Exit Function
ErrorHandler:
    'handle the error
    lRet = Err.Number
    
    m_oErr.HandleError lRet, Err.Description, p_bDisplayMessage:=True
    GoTo Exit_Point
End Function

Private Sub treMTS_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Me.MousePointer = vbSizeWE Then
        Me.MousePointer = vbDefault
    End If
End Sub
Private Function GetSettings() As Long
    Dim lRet As Long
    Dim oReg As CRegistry
    Dim lTop As Long
    Dim lLeft As Long
    Dim lHeight As Long
    Dim lWidth As Long
    Dim colValues As Collection
    Dim i As Long
    
    Const PROC_NAME As String = "GetSettings"
    On Error GoTo ErrorHandler
    m_oErr.AddProcedureName PROC_NAME
    
    'create the registry component
    Set oReg = New CRegistry
    
    lRet = oReg.OpenRegistry(HKEY_LOCAL_MACHINE, "SOFTWARE\WOW\Client\Screen")
    
    If lRet Then
        'get the values
        Set colValues = oReg.GetAllValues
                
    Else
        oReg.OpenRegistry HKEY_LOCAL_MACHINE, ""
        
        oReg.CreateDirectory "SOFTWARE\WOW\Client\Screen"
        
        'create the defaults
        oReg.CreateValue "Top", 0, REG_SZ
        oReg.CreateValue "Left", 0, REG_SZ
        oReg.CreateValue "Width", Screen.Width, REG_SZ
        oReg.CreateValue "Height", Screen.Height, REG_SZ
        oReg.CreateValue "Splitter", 4000, REG_SZ
        oReg.CreateValue "WindowState", vbNormal, REG_SZ
        
        oReg.OpenRegistry HKEY_LOCAL_MACHINE, "SOFTWARE\WOW\Client\Screen"
        Set colValues = oReg.GetAllValues
        
    End If
    
    'move the form
    frmMTSAdmin.Move colValues("Left"), colValues("Top"), colValues("Width"), colValues("Height")
    
    'set the splitter location
    m_lLeft = colValues("Splitter")
    
    'set the windowstate
    frmMTSAdmin.WindowState = colValues("WindowState")
    
    'get the computer list
    lRet = oReg.OpenRegistry(HKEY_LOCAL_MACHINE, "SOFTWARE\WOW\Client\Computers")
    
    If lRet Then
        Set colValues = oReg.GetAllValues
        
    Else
        oReg.OpenRegistry HKEY_LOCAL_MACHINE, ""
        oReg.CreateDirectory "SOFTWARE\WOW\Client\Computers"
        oReg.OpenRegistry HKEY_LOCAL_MACHINE, "SOFTWARE\WOW\Client\Computers"
        Set colValues = oReg.GetAllValues
    End If
    
    For i = 1 To colValues.Count
        AddComputer colValues.Item(i)
    Next i
Exit_Point:
    'return the results
    GetSettings = lRet
    
    'clean up
    oReg.CloseRegistry
    Set oReg = Nothing
    Set colValues = Nothing
    
    'remove us from the error stack
    m_oErr.RemoveProcedureName PROC_NAME
    Exit Function
ErrorHandler:
    'handle the error
    lRet = Err.Number
    
    m_oErr.HandleError lRet, Err.Description
    GoTo Exit_Point
End Function
Private Function AddComputer(ByVal p_sComputerName As String) As Long
    Const PROC_NAME As String = "AddComputer"
    Dim lRet As Long
    Dim oNode As MSComctlLib.Node
    
    On Error GoTo ErrorHandler
    
    m_oErr.AddProcedureName PROC_NAME
    
    If m_oRoot.Expanded = False Then
        m_oRoot.Expanded = True
    End If
    
    m_bProcessExpand = False
    
    'add the new node
    Set oNode = treMTS.Nodes.Add
    oNode.Text = p_sComputerName
    oNode.Key = "Computer" & m_oRoot.children + 1
    Set oNode.Parent = m_oRoot
    oNode.Image = "RemoteComputer"
    
    'add a blank node
    AddBlankNode oNode
    
    'make sure it's selected
    oNode.Selected = True
    If p_sComputerName = "" Then
        treMTS.StartLabelEdit
    Else
        m_oRoot.Sorted = True
    End If
    
Exit_Point:
    'return the results
    
    'clean up
    Set oNode = Nothing
    m_bProcessExpand = True
    
    'remove us from the error stack
    m_oErr.RemoveProcedureName PROC_NAME
    Exit Function
ErrorHandler:
    'handle the error
    lRet = Err.Number
    
    m_oErr.HandleError lRet, Err.Description, p_bDisplayMessage:=True
    GoTo Exit_Point

End Function
Private Function SaveSettings() As Long
    Dim lRet As Long
    Dim oReg As CRegistry
    Dim lTop As Long
    Dim lLeft As Long
    Dim lHeight As Long
    Dim lWidth As Long
    Dim oNode As MSComctlLib.Node
    
    Dim i As Long
    
    Const PROC_NAME As String = "SaveSettings"
    On Error GoTo ErrorHandler
    m_oErr.AddProcedureName PROC_NAME
    
    'create the registry component
    Set oReg = New CRegistry
    
    lRet = oReg.OpenRegistry(HKEY_LOCAL_MACHINE, "SOFTWARE\WOW\Client\Screen")
    
    If lRet Then
        'set the values
        If frmMTSAdmin.WindowState = vbMaximized Or frmMTSAdmin.WindowState = vbMinimized Then
            oReg.CreateValue "WindowState", frmMTSAdmin.WindowState, REG_SZ
        Else
            oReg.CreateValue "Top", frmMTSAdmin.Top, REG_SZ
            oReg.CreateValue "Left", frmMTSAdmin.Left, REG_SZ
            oReg.CreateValue "Width", frmMTSAdmin.Width, REG_SZ
            oReg.CreateValue "Height", frmMTSAdmin.Height, REG_SZ
            oReg.CreateValue "Splitter", m_lLeft, REG_SZ
            oReg.CreateValue "WindowState", frmMTSAdmin.WindowState, REG_SZ
        End If
    End If
    
    
    'get the computer list
    lRet = oReg.OpenRegistry(HKEY_LOCAL_MACHINE, "SOFTWARE\WOW\Client\Computers")
    
    If lRet Then
        Set oNode = m_oRoot.Child
        'skip the my computer node
        While Not oNode Is Nothing
            If oNode.Text <> "My Computer" Then
                oReg.CreateValue oNode.Text, oNode.Text, REG_SZ
                Set oNode = oNode.Next
            Else
                Set oNode = oNode.Next
            End If
        Wend
        
    End If
    
Exit_Point:
    'return the results
    SaveSettings = lRet
    
    'clean up
    oReg.CloseRegistry
    Set oReg = Nothing
    
    'remove us from the error stack
    m_oErr.RemoveProcedureName PROC_NAME
    Exit Function
ErrorHandler:
    'handle the error
    lRet = Err.Number
    
    m_oErr.HandleError lRet, Err.Description
    GoTo Exit_Point
End Function

Private Sub treMTS_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim oParent As IXMLDOMNode
    Dim oList As IXMLDOMNodeList
    Dim j As Long
    Dim i As Long
    Dim lRet As Long
    Const PROC_NAME As String = "treMTS_NodeClick"
    On Error GoTo ErrorHandler
    m_oErr.AddProcedureName PROC_NAME
    Me.MousePointer = vbHourglass
    LockWindowUpdate grdDetails.hWnd
    Select Case Left$(Node.Key, 8)
        Case "Package:"
            'get the components for this package
            If Node.Parent.Parent.Text = "My Computer" Then
                lRet = GetComponents("", Node, oList)
            Else
                lRet = GetComponents(Node.Parent.Parent.Text, Node, oList)
            End If
            grdDetails.Rows = oList.length + 1
            grdDetails.Cols = oList.Item(0).childNodes.length + 1
            grdDetails.Row = 0
            For j = 0 To oList.Item(0).childNodes.length - 1
                grdDetails.Col = j + 1
                grdDetails.Text = oList.Item(0).childNodes(j).nodeName
            Next j
            For i = 0 To oList.length - 1
                Set oParent = oList.Item(i)
                'set the image
                grdDetails.Col = 0
                grdDetails.Row = i + 1
                grdDetails.ColWidth(0) = 300
                Set grdDetails.CellPicture = ImageList16.ListImages("Component").Picture
                For j = 0 To oParent.childNodes.length - 1
                    grdDetails.Col = j + 1
                    grdDetails.Text = oParent.childNodes(j).Text
                Next j
            Next i
        Case "PackFold"
            'get the components for this package
            If Node.Parent.Text = "My Computer" Then
                lRet = GetPackages("", Node, oList)
            Else
                lRet = GetPackages(Node.Parent.Text, Node, oList)
            End If
            grdDetails.Rows = oList.length + 1
            grdDetails.Cols = oList.Item(0).childNodes.length + 1
            grdDetails.Row = 0
            For j = 0 To oList.Item(0).childNodes.length - 1
                grdDetails.Col = j + 1
                grdDetails.Text = oList.Item(0).childNodes(j).nodeName
            Next j
            For i = 0 To oList.length - 1
                Set oParent = oList.Item(i)
                'set the image
                grdDetails.Col = 0
                grdDetails.Row = i + 1
                grdDetails.ColWidth(0) = 300
                Set grdDetails.CellPicture = ImageList16.ListImages("Package").Picture
                For j = 0 To oParent.childNodes.length - 1
                    grdDetails.Col = j + 1
                    grdDetails.Text = oParent.childNodes(j).Text
                Next j
            Next i
        
        Case Else
    End Select
    
    grdDetails.Col = 2
    grdDetails.ColSel = 1
    grdDetails.Sort = flexSortStringNoCaseAscending
Exit_Point:
    'return the results
    
    'clean up
    Set oList = Nothing
    Set oParent = Nothing
    Me.MousePointer = vbDefault
    LockWindowUpdate 0
    
    'remove us from the error stack
    m_oErr.RemoveProcedureName PROC_NAME
    Exit Sub
ErrorHandler:
    'handle the error
    lRet = Err.Number
    
    m_oErr.HandleError lRet, Err.Description
    GoTo Exit_Point
End Sub
