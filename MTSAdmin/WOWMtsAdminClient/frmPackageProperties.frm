VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmPackageProperties 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Package Properties"
   ClientHeight    =   4995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   5895
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   315
      Left            =   2640
      TabIndex        =   22
      Top             =   4620
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   3720
      TabIndex        =   23
      Top             =   4620
      Width           =   1035
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   315
      Left            =   4800
      TabIndex        =   24
      Top             =   4620
      Width           =   1035
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4515
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5835
      _ExtentX        =   10292
      _ExtentY        =   7964
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "frmPackageProperties.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Line2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Line1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "imgComponent"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtPackageID"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtName"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtDescription"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Security"
      TabPicture(1)   =   "frmPackageProperties.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraAuthorization"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Identity"
      TabPicture(2)   =   "frmPackageProperties.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label7"
      Tab(2).Control(1)=   "fraSecurity"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Activation"
      TabPicture(3)   =   "frmPackageProperties.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fraActivation"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Advanced"
      TabPicture(4)   =   "frmPackageProperties.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "fraShutDown"
      Tab(4).Control(1)=   "fraPermission"
      Tab(4).ControlCount=   2
      Begin VB.Frame fraPermission 
         Caption         =   "Permission"
         Height          =   1095
         Left            =   -74880
         TabIndex        =   19
         Top             =   1860
         Width           =   5535
         Begin VB.CheckBox chkDisableDeletion 
            Caption         =   "Disable deletion"
            Height          =   315
            Left            =   240
            TabIndex        =   21
            Top             =   660
            Width           =   1575
         End
         Begin VB.CheckBox chkDisableChanges 
            Caption         =   "Disable changes"
            Height          =   255
            Left            =   240
            TabIndex        =   20
            Top             =   300
            Width           =   1635
         End
      End
      Begin VB.Frame fraShutDown 
         Caption         =   "Server Process Shutdown"
         Height          =   1095
         Left            =   -74880
         TabIndex        =   15
         Top             =   600
         Width           =   5535
         Begin VB.TextBox txtMinutes 
            Height          =   315
            Left            =   2460
            TabIndex        =   18
            Top             =   600
            Width           =   2895
         End
         Begin VB.OptionButton optShutdown 
            Caption         =   "Minutes until shutdown:"
            Height          =   255
            Left            =   240
            TabIndex        =   17
            Top             =   660
            Width           =   2055
         End
         Begin VB.OptionButton optRunForever 
            Caption         =   "Leave running when idle"
            Height          =   195
            Left            =   240
            TabIndex        =   16
            Top             =   300
            Width           =   2415
         End
      End
      Begin VB.Frame fraActivation 
         Caption         =   "Activation Mode"
         Height          =   1275
         Left            =   -74880
         TabIndex        =   12
         Top             =   540
         Width           =   5535
         Begin VB.OptionButton optLocal 
            Caption         =   "Library Package"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Tag             =   "Local"
            Top             =   240
            Width           =   1455
         End
         Begin VB.OptionButton optServer 
            Caption         =   "Server Package"
            Height          =   195
            Left            =   120
            TabIndex        =   14
            Tag             =   "InProc"
            Top             =   780
            Value           =   -1  'True
            Width           =   1455
         End
         Begin VB.Label Label9 
            Caption         =   "Components will be activated in the creator's process."
            Height          =   255
            Left            =   420
            TabIndex        =   33
            Top             =   480
            Width           =   3855
         End
         Begin VB.Label Label8 
            Caption         =   "Components will be activated in a dedicated server process."
            Height          =   195
            Left            =   420
            TabIndex        =   32
            Top             =   1020
            Width           =   4275
         End
      End
      Begin VB.Frame fraSecurity 
         Caption         =   "Security"
         Height          =   1815
         Left            =   -74880
         TabIndex        =   6
         Top             =   780
         Width           =   5535
         Begin VB.OptionButton optInteractive 
            Caption         =   "Interactive User--The currently logged on user."
            Height          =   195
            Left            =   120
            TabIndex        =   7
            Top             =   300
            Value           =   -1  'True
            Width           =   4875
         End
         Begin VB.OptionButton optSpecificUser 
            Caption         =   "This User:"
            Height          =   195
            Left            =   120
            TabIndex        =   8
            Top             =   540
            Width           =   1815
         End
         Begin VB.TextBox txtUser 
            Height          =   315
            Left            =   1560
            TabIndex        =   9
            Top             =   780
            Width           =   1995
         End
         Begin VB.TextBox txtPassword 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   1560
            PasswordChar    =   "*"
            TabIndex        =   10
            Top             =   1140
            Width           =   1995
         End
         Begin VB.TextBox txtConfirm 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   1560
            PasswordChar    =   "*"
            TabIndex        =   11
            Top             =   1440
            Width           =   1995
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "User Name:"
            Height          =   255
            Left            =   180
            TabIndex        =   30
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "Password:"
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   1200
            Width           =   1395
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Confirm Password:"
            Height          =   195
            Left            =   120
            TabIndex        =   28
            Top             =   1500
            Width           =   1395
         End
      End
      Begin VB.Frame fraAuthorization 
         Caption         =   "Authorization"
         Height          =   1695
         Left            =   -74880
         TabIndex        =   27
         Top             =   480
         Width           =   5535
         Begin VB.ComboBox comAuthenticationLevel 
            Height          =   315
            ItemData        =   "frmPackageProperties.frx":008C
            Left            =   120
            List            =   "frmPackageProperties.frx":00A2
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   1140
            Width           =   5295
         End
         Begin VB.CheckBox chkAuthorization 
            Caption         =   "Enforce authorization checks for this package"
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   300
            Width           =   5175
         End
         Begin VB.Label Label2 
            Caption         =   "Authentication level for calls:"
            Height          =   255
            Left            =   180
            TabIndex        =   34
            Top             =   900
            Width           =   2235
         End
      End
      Begin VB.TextBox txtDescription 
         Height          =   1275
         Left            =   180
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   1980
         Width           =   5535
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   960
         TabIndex        =   1
         Text            =   "Package Name"
         Top             =   840
         Width           =   4755
      End
      Begin VB.TextBox txtPackageID 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   3
         Text            =   "Package ID"
         Top             =   3360
         Width           =   4635
      End
      Begin VB.Label Label7 
         Caption         =   "The package will run under the following user context:"
         Height          =   255
         Left            =   -74820
         TabIndex        =   31
         Top             =   480
         Width           =   4635
      End
      Begin VB.Image imgComponent 
         Height          =   480
         Left            =   300
         Picture         =   "frmPackageProperties.frx":00E5
         Top             =   720
         Width           =   480
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000011&
         X1              =   120
         X2              =   5700
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Line Line2 
         BorderColor     =   &H8000000E&
         X1              =   120
         X2              =   5700
         Y1              =   1455
         Y2              =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Description:"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   1680
         Width           =   915
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Package:"
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   3360
         Width           =   915
      End
   End
   Begin VB.Label lblDebug 
      Height          =   375
      Left            =   300
      TabIndex        =   35
      Top             =   4620
      Width           =   915
   End
End
Attribute VB_Name = "frmPackageProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_oDOM As DOMDocument30
Private m_sServerName As String
Private m_bDirty As Boolean
Private m_bLoading As Boolean
Private m_sTransactions As String
Private m_bUserChanged As Boolean
Private m_bChanged As Boolean
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

Private Sub chkDisableChanges_Click()
    Dirty = True
End Sub

Private Sub chkDisableDeletion_Click()
    Dirty = True
End Sub

Private Sub cmdApply_Click()
    SetProperties
End Sub

Private Sub cmdCancel_Click()
    Unload frmPackageProperties
End Sub

Private Sub cmdOK_Click()
    If m_bDirty = True Then
        SetProperties
    End If
    Unload frmPackageProperties
End Sub

Private Sub comAuthenticationLevel_Click()
    Dirty = True
End Sub

Private Sub Form_Activate()
    Dim i As Long
    Dim lRet As Long

    Const PROC_NAME As String = "Form_Activate"
    On Error GoTo ErrorHandler
    m_oErr.AddProcedureName PROC_NAME
    
    'place the data
    m_bLoading = True
    txtName.Text = NodeValue("//Name")
    frmPackageProperties.Caption = NodeValue("//Name") & " Package Properties"
    txtDescription.Text = NodeValue("//Description")
    txtPackageID.Text = NodeValue("//ID")
    If NodeValue("//SecurityEnabled") = "Y" Then
        chkAuthorization.Value = 1
    Else
        chkAuthorization.Value = 0
    End If
    For i = 0 To comAuthenticationLevel.ListCount - 1
        If comAuthenticationLevel.ItemData(i) = NodeValue("//Authentication") Then
            comAuthenticationLevel.ListIndex = i
            Exit For
        End If
    Next i
    If comAuthenticationLevel.ListIndex = -1 Then comAuthenticationLevel.ListIndex = 0
    If NodeValue("//Identity") = "Interactive User" Then
        optInteractive.Value = 1
    Else
        optSpecificUser.Value = 1
        txtUser.Text = NodeValue("//Identity")
    End If
    If NodeValue("//Activation") = "Inproc" Then
        optLocal.Value = 1
    Else
        optServer.Value = 1
    End If
    If NodeValue("//RunForever") = "Y" Then
        optRunForever.Value = True
    Else
        optShutdown.Value = True
        txtMinutes.Text = NodeValue("//ShutdownAfter")
    End If
    If NodeValue("//Changeable") = "Y" Then
        chkDisableChanges = 0
    Else
        chkDisableChanges = 1
        txtName.Enabled = False
        txtDescription.Enabled = False
        fraAuthorization.Enabled = False
        fraSecurity.Enabled = False
        fraActivation.Enabled = False
        fraShutDown.Enabled = False
        chkDisableDeletion.Enabled = False
    End If
    If NodeValue("//Deleteable") = "Y" Then
        chkDisableDeletion = 0
    Else
        chkDisableDeletion = 1
    End If
    
'    If NodeValue("//IsSystem") = "Y" Then
'        txtDescription.Enabled = False
'        fraTransactions.Enabled = False
'        fraAuthorization.Enabled = False
'    End If
    
Exit_Point:
    'return the results
    
    'clean up
    m_bLoading = False
    
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
    lRet = MsgBox(m_oDOM.xml & vbCrLf & "Would you like to save this output?", vbYesNo + vbInformation, "Package Properties")
    If lRet = vbYes Then
        m_oDOM.save InputBox$("Please enter the file name:", "File Name", "c:\PackageProperties.xml")
    End If

End Sub

Private Sub optInteractive_Click()
    Dirty = True
End Sub

Private Sub optLocal_Click()
    Dirty = True
    fraSecurity.Enabled = False
    fraAuthorization.Enabled = False
    fraShutDown.Enabled = False
    
End Sub

Private Sub optRunForever_Click()
    Dirty = True
End Sub

Private Sub optServer_Click()
    Dirty = True
    fraSecurity.Enabled = True
    fraAuthorization.Enabled = True
    fraShutDown.Enabled = True
    
End Sub

Private Sub optShutdown_Click()
    Dirty = True
End Sub

Private Sub optSpecificUser_Click()
    Dirty = True
End Sub

Private Sub txtConfirm_Change()
    Dirty = True
End Sub

Private Sub txtDescription_Change()
    Dirty = True
End Sub
Private Function SetProperties() As Long
    Dim lRet As Long
    Dim sXML As String
    Dim oDOM As DOMDocument30
    Dim oMTSAdmin As WOWMTSAdmin.CMtsAdmin
    Dim oPackage As IXMLDOMNode
    Dim oNode As IXMLDOMNode
    
    Const PROC_NAME As String = "SetProperties"
    On Error GoTo ErrorHandler
    m_oErr.AddProcedureName PROC_NAME
    
    Me.MousePointer = vbHourglass
    
    'check the password first
    If optSpecificUser.Value = 1 Then
        If txtPassword.Text = "" And m_bUserChanged = True Then
            Call MsgBox("You didn't supply a password.", vbOKOnly + vbExclamation, "Password Error")
            SetProperties = -1
            Exit Function
        End If
        If StrComp(txtPassword.Text, txtConfirm.Text) <> 0 Then
            Call MsgBox("Your passwords didn't match", vbOKOnly + vbExclamation, "Password Error")
            SetProperties = -1
            Exit Function
        End If
    End If
    
    'init the dom
    InitDOM oDOM
    
    'load the default xml string
    oDOM.loadXML XMLSTRING
            
    'get the package node
    Set oPackage = oDOM.selectSingleNode("/Root/Package")
    
    'set the new node
    Set oNode = m_oDOM.selectSingleNode("/Package")
    
    'replace the xml
    oDOM.selectSingleNode("/Root").replaceChild oNode, oPackage
    
    Set oPackage = oDOM.selectSingleNode("/Root/Package")
    
    'remove the identity node--we'll add it again later if we need it
    If NodeExists("Identity", oPackage) Then
        Set oNode = oPackage.selectSingleNode("Identity")
        oPackage.removeChild oNode
    End If
    
    'set the node values
    SetNodeValue "/Root/ServerName", m_sServerName, oDOM
    SetNodeValue "/Root/Package/Name", txtName.Text, oDOM
    SetNodeValue "/Root/Package/Description", txtDescription.Text, oDOM
    If optLocal.Value = True Then
        SetNodeValue "/Root/Package/Activation", "InProc", oDOM
        Set oNode = oPackage.selectSingleNode("SecurityEnabled")
        oPackage.removeChild oNode
        Set oNode = oPackage.selectSingleNode("Authentication")
        oPackage.removeChild oNode
        Set oNode = oPackage.selectSingleNode("RunForever")
        oPackage.removeChild oNode
        Set oNode = oPackage.selectSingleNode("ShutdownAfter")
        oPackage.removeChild oNode
    Else
        SetNodeValue "/Root/Package/Activation", "Local", oDOM
        If optInteractive.Value = True Then
            AddElement oPackage, "Identity", "Interactive User"
        Else
            If m_bUserChanged = True Then
                AddElement oPackage, "Identity", txtUser.Text
                AddElement oPackage, "Password", txtPassword.Text
            End If
        End If
        If chkAuthorization.Value = 0 Then
            SetNodeValue "/Root/Package/SecurityEnabled", "N", oDOM
        Else
            SetNodeValue "/Root/Package/SecurityEnabled", "Y", oDOM
        End If
        SetNodeValue "/Root/Package/Authentication", comAuthenticationLevel.ItemData(comAuthenticationLevel.ListIndex), oDOM
        If optRunForever.Value = True Then
            SetNodeValue "/Root/Package/RunForever", "Y", oDOM
        Else
            SetNodeValue "/Root/Package/RunForever", "N", oDOM
            SetNodeValue "/Root/Package/ShutdownAfter", txtMinutes.Text, oDOM
        End If
    End If
    If chkDisableChanges.Value = 1 Then
        SetNodeValue "/Root/Package/Changeable", "N", oDOM
    Else
        SetNodeValue "/Root/Package/Changeable", "Y", oDOM
    End If
    If chkDisableDeletion.Value = 1 Then
        SetNodeValue "/Root/Package/Deletable", "N", oDOM
    Else
        SetNodeValue "/Root/Package/Deletable", "Y", oDOM
    End If
    
    'Load the changes into the dom
    m_oDOM.loadXML oDOM.selectSingleNode("/Root/Package").xml
    
    'create an instance of the mts admin on the computer supplied
    If oDOM.selectSingleNode("//ServerName").Text = "" Then
        Set oMTSAdmin = CreateObject("WOWMTSAdmin.CMtsAdmin")
    Else
        Set oMTSAdmin = CreateObject("WOWMTSAdmin.CMtsAdmin", oDOM.selectSingleNode("//ServerName").Text)
    End If
    
    'set the package properties
    lRet = oMTSAdmin.SetPackageProperties(oDOM.xml, sXML)
                
Exit_Point:
    'return the results
    SetProperties = lRet
        
    'clean up
    Set oMTSAdmin = Nothing
    Set oNode = Nothing
    Set oDOM = Nothing
    Set oPackage = Nothing
    Dirty = False
    m_bUserChanged = False
    m_bChanged = True
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
    Dim oNode As IXMLDOMNode
    Dim oParent As IXMLDOMNode
    Dim lRet As Long

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


Private Sub txtMinutes_Change()
    Dirty = True
End Sub

Private Sub txtName_Change()
    Dirty = True
End Sub

Private Sub txtPassword_Change()
    Dirty = True
    If m_bLoading = False Then
        m_bUserChanged = True
    End If
End Sub

Private Sub txtUser_Change()
    Dirty = True
    If m_bLoading = False Then
        m_bUserChanged = True
    End If
End Sub
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
    
    m_oErr.HandleError lRet, Err.Description, p_bDisplayMessage:=True
    GoTo Exit_Point
End Function

