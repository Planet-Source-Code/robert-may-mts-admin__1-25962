VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmComponentProperties 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Component Properties"
   ClientHeight    =   4995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5865
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   5865
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   315
      Left            =   4800
      TabIndex        =   17
      Top             =   4620
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   3720
      TabIndex        =   16
      Top             =   4620
      Width           =   1035
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   315
      Left            =   2640
      TabIndex        =   15
      Top             =   4620
      Width           =   1035
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4515
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   5835
      _ExtentX        =   10292
      _ExtentY        =   7964
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "frmComponentProperties.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "imgComponent"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Line1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Line2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label3"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label4"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtDescription"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtName"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtDLL"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtClSID"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtPackageID"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "Transactions"
      TabPicture(1)   =   "frmComponentProperties.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraTransactions"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Security"
      TabPicture(2)   =   "frmComponentProperties.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraAuthorization"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Concurrency"
      TabPicture(3)   =   "frmComponentProperties.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fraThreadingModel"
      Tab(3).ControlCount=   1
      Begin VB.Frame fraThreadingModel 
         Caption         =   "Threading Model"
         Height          =   795
         Left            =   -74880
         TabIndex        =   14
         Top             =   480
         Width           =   5535
         Begin VB.TextBox txtThreadingModel 
            BackColor       =   &H8000000A&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   180
            Locked          =   -1  'True
            TabIndex        =   13
            Text            =   "Threading Model"
            Top             =   300
            Width           =   4875
         End
      End
      Begin VB.Frame fraAuthorization 
         Caption         =   "Authorization"
         Height          =   795
         Left            =   -74880
         TabIndex        =   11
         Top             =   480
         Width           =   5535
         Begin VB.CheckBox chkAuthorization 
            Caption         =   "Enforce component level access checks"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   300
            Width           =   3315
         End
      End
      Begin VB.Frame fraTransactions 
         Caption         =   "Transaction Support"
         Height          =   1515
         Left            =   -74880
         TabIndex        =   5
         Top             =   480
         Width           =   5535
         Begin VB.OptionButton optRequiresNew 
            Caption         =   "Requires New"
            Height          =   195
            Left            =   120
            TabIndex        =   10
            Tag             =   "Requires New"
            Top             =   1200
            Width           =   4815
         End
         Begin VB.OptionButton optRequired 
            Caption         =   "Required"
            Height          =   195
            Left            =   120
            TabIndex        =   9
            Tag             =   "Required"
            Top             =   960
            Width           =   4635
         End
         Begin VB.OptionButton optSupported 
            Caption         =   "Supported"
            Height          =   195
            Left            =   120
            TabIndex        =   8
            Tag             =   "Supported"
            Top             =   720
            Width           =   4755
         End
         Begin VB.OptionButton optNotSupported 
            Caption         =   "Not Supported"
            Height          =   195
            Left            =   120
            TabIndex        =   7
            Tag             =   "Not Supported"
            Top             =   480
            Width           =   4635
         End
         Begin VB.OptionButton optDisabled 
            Caption         =   "Disabled"
            Height          =   195
            Left            =   120
            TabIndex        =   6
            Tag             =   "Ignored"
            Top             =   240
            Width           =   3795
         End
      End
      Begin VB.TextBox txtPackageID 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   4
         Text            =   "Package ID"
         Top             =   4140
         Width           =   4635
      End
      Begin VB.TextBox txtClSID 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   3
         Text            =   "Class ID"
         Top             =   3780
         Width           =   4635
      End
      Begin VB.TextBox txtDLL 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   2
         Text            =   "DLL"
         Top             =   3420
         Width           =   4635
      End
      Begin VB.TextBox txtName 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   0
         Text            =   "Component Name"
         Top             =   840
         Width           =   4695
      End
      Begin VB.TextBox txtDescription 
         Height          =   1275
         Left            =   180
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   1980
         Width           =   5535
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Package:"
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   4140
         Width           =   915
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "CLSID:"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   3780
         Width           =   915
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "DLL:"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   3420
         Width           =   915
      End
      Begin VB.Label Label1 
         Caption         =   "Description:"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   1680
         Width           =   915
      End
      Begin VB.Line Line2 
         BorderColor     =   &H8000000E&
         X1              =   120
         X2              =   5700
         Y1              =   1455
         Y2              =   1455
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000011&
         X1              =   120
         X2              =   5700
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Image imgComponent 
         Height          =   480
         Left            =   300
         Picture         =   "frmComponentProperties.frx":0070
         Top             =   720
         Width           =   480
      End
   End
   Begin VB.Label lblDebug 
      Height          =   315
      Left            =   300
      TabIndex        =   23
      Top             =   4620
      Width           =   555
   End
End
Attribute VB_Name = "frmComponentProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_oDOM As DOMDocument30
Private m_sServerName As String
Private m_bDirty As Boolean
Private m_bChanged As Boolean
Private m_bLoading As Boolean
Private m_sTransactions As String
Private m_oErr As CError
Public Property Set DOM(ByVal r_oDOM As DOMDocument30)
    Set m_oDOM = r_oDOM
End Property
Public Property Let ServerName(ByVal p_sServer As String)
    m_sServerName = p_sServer
End Property
Private Sub chkAuthorization_Click()
    Dirty = True
End Sub

Private Sub cmdApply_Click()
    SetProperties
End Sub

Private Sub cmdCancel_Click()
    Unload frmComponentProperties
End Sub

Private Sub cmdOK_Click()
    If m_bDirty = True Then
        SetProperties
    End If
    Unload frmComponentProperties
End Sub

Private Sub Form_Activate()
    Dim lRet As Long
    Const PROC_NAME As String = "Form_Activate"
    On Error GoTo ErrorHandler
    m_oErr.AddProcedureName PROC_NAME
    'place the data
    m_bLoading = True
    txtName.Text = NodeValue("//ProgID")
    frmComponentProperties.Caption = NodeValue("//ProgID") & " Component Properties"
    txtDescription.Text = NodeValue("//Description")
    txtDLL.Text = NodeValue("//DLL")
    txtClSID.Text = NodeValue("//CLSID")
    txtPackageID.Text = NodeValue("//PackageID")
    txtThreadingModel.Text = NodeValue("//ThreadingModel")
    If NodeValue("//SecurityEnabled") = "Y" Then
        chkAuthorization.Value = 1
    Else
        chkAuthorization.Value = 0
    End If
    Select Case UCase$(NodeValue("//Transaction"))
        Case "IGNORED"
            optDisabled.Value = True
        Case "NOT SUPPORTED"
            optNotSupported.Value = True
        Case "SUPPORTED"
            optSupported = True
        Case "REQUIRED"
            optRequired = True
        Case "REQUIRES NEW"
            optRequiresNew = True
    End Select
    If NodeValue("//IsSystem") = "Y" Then
        txtDescription.Enabled = False
        fraTransactions.Enabled = False
        fraAuthorization.Enabled = False
    End If
    
    m_bLoading = False
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

Private Sub Form_Load()
    Set m_oErr = New CError
    Dirty = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim oNode As MSComctlLib.Node
    Set m_oDOM = Nothing
    If m_bChanged = True Then
        Set oNode = frmMTSAdmin.CurrentNode
        frmMTSAdmin.RefreshTree oNode.Parent
    End If
    Set oNode = Nothing
    Set m_oErr = Nothing
End Sub
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
    
    m_oErr.HandleError lRet, Err.Description, p_bDisplayMessage:=True
    GoTo Exit_Point
End Function
Private Property Let Dirty(ByVal p_bDirty As Boolean)
    If m_bLoading = False Then
        m_bDirty = p_bDirty
        cmdApply.Enabled = p_bDirty
    End If
End Property

Private Sub lblDebug_DblClick()
    Dim lRet As Long
    On Error Resume Next
    lRet = MsgBox(m_oDOM.xml & vbCrLf & "Would you like to save this output?", vbYesNo + vbInformation, "Component Properties")
    If lRet = vbYes Then
        m_oDOM.save InputBox$("Please enter the file name:", "File Name", "c:\ComponentProperties.xml")
    End If

End Sub

Private Sub optDisabled_Click()
    Dirty = True
    m_sTransactions = optDisabled.Tag
End Sub

Private Sub optNotSupported_Click()
    Dirty = True
    m_sTransactions = optNotSupported.Tag
End Sub

Private Sub optRequired_Click()
    Dirty = True
    m_sTransactions = optRequired.Tag
End Sub

Private Sub optRequiresNew_Click()
    Dirty = True
    m_sTransactions = optRequiresNew.Tag
End Sub

Private Sub optSupported_Click()
    Dirty = True
    m_sTransactions = optSupported.Tag
End Sub

Private Sub txtDescription_Change()
    Dirty = True
End Sub
Private Function SetProperties() As Long
    Dim lRet As Long
    Dim sXML As String
    Dim oDOM As DOMDocument30
    Dim oMTSAdmin As WOWMTSAdmin.CMtsAdmin
    Dim oComponent As IXMLDOMNode
    Dim oNode As IXMLDOMNode
    Const PROC_NAME As String = "SetProperties"
    
    On Error GoTo ErrorHandler
    m_oErr.AddProcedureName PROC_NAME
    Me.MousePointer = vbHourglass
    
    'init the dom
    InitDOM oDOM
    
    'load the default xml string
    oDOM.loadXML XMLSTRING
            
    'get the package node
    Set oComponent = oDOM.selectSingleNode("/Root/Package/Component")
    
    'set the new node
    Set oNode = m_oDOM.selectSingleNode("/Component")
    
    'replace the xml
    oDOM.selectSingleNode("/Root/Package").replaceChild oNode, oComponent
    
    'set the noe values
    SetNodeValue "/Root/Package/ID", oNode.selectSingleNode("PackageID").Text, oDOM
    SetNodeValue "/Root/ServerName", m_sServerName, oDOM
    SetNodeValue "/Root/Package/Component/Transaction", m_sTransactions, oDOM
    If chkAuthorization.Value = 0 Then
        SetNodeValue "/Root/Package/Component/SecurityEnabled", "N", oDOM
    Else
        SetNodeValue "/Root/Package/Component/SecurityEnabled", "Y", oDOM
    End If
    SetNodeValue "/Root/Package/Component/Description", txtDescription, oDOM
    
    'load the updates
    m_oDOM.loadXML oDOM.selectSingleNode("/Root/Package/Component").xml
            
    'create an instance of the mts admin on the computer supplied
    If oDOM.selectSingleNode("//ServerName").Text = "" Then
        Set oMTSAdmin = CreateObject("WOWMTSAdmin.CMtsAdmin")
    Else
        Set oMTSAdmin = CreateObject("WOWMTSAdmin.CMtsAdmin", oDOM.selectSingleNode("//ServerName").Text)
    End If
    
    'get the packages
    lRet = oMTSAdmin.SetComponentProperties(oDOM.xml, sXML)
            
    'clean up
    Dirty = False
    m_bChanged = True
    
Exit_Point:
    'return the results
    SetProperties = lRet
    
    'clean up
    Set oMTSAdmin = Nothing
    Set oNode = Nothing
    Set oDOM = Nothing
    Set oComponent = Nothing
    Me.MousePointer = vbDefault
    
    'remove us from the error stack
    m_oErr.RemoveProcedureName PROC_NAME
    Exit Function
ErrorHandler:
    'handle the error
    lRet = Err.Number
    
    m_oErr.HandleError lRet, Err.Description, p_bDisplayMessage:=True
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
    
    m_oErr.HandleError lRet, Err.Description, p_bDisplayMessage:=True
    GoTo Exit_Point
End Function
Private Function SetNodeValue(ByVal p_sNodeName As String, ByVal p_sValue As String, Optional ByRef p_oParent As IXMLDOMNode) As Long
    Dim lRet As Long
    Dim oNode As IXMLDOMNode
    Dim oParent As IXMLDOMNode
    Const PROC_NAME As String = "SetNodeValue"
    
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
    Set oParent = Nothing
    
    'remove us from the error stack
    m_oErr.RemoveProcedureName PROC_NAME
    Exit Function
ErrorHandler:
    'handle the error
    lRet = Err.Number
    
    m_oErr.HandleError lRet, Err.Description, p_bDisplayMessage:=True
    GoTo Exit_Point
End Function

