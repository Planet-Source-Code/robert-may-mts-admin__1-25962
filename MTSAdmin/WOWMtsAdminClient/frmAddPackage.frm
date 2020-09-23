VERSION 5.00
Begin VB.Form frmAddPackage 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add New Package"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5265
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   5265
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add Package"
      Height          =   375
      Left            =   1680
      TabIndex        =   9
      Top             =   3960
      Width           =   1815
   End
   Begin VB.Frame fraSecurity 
      Caption         =   "Security"
      Height          =   1815
      Left            =   0
      TabIndex        =   3
      Top             =   2040
      Width           =   5235
      Begin VB.TextBox txtConfirm 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1560
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   1440
         Width           =   1995
      End
      Begin VB.TextBox txtPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1560
         PasswordChar    =   "*"
         TabIndex        =   7
         Top             =   1140
         Width           =   1995
      End
      Begin VB.TextBox txtUser 
         Height          =   315
         Left            =   1560
         TabIndex        =   6
         Top             =   780
         Width           =   1995
      End
      Begin VB.OptionButton optSpecificUser 
         Caption         =   "This User:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   540
         Width           =   1815
      End
      Begin VB.OptionButton optInteractive 
         Caption         =   "Interactive User--The currently logged on user."
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   300
         Value           =   -1  'True
         Width           =   4875
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Confirm Password:"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   1500
         Width           =   1395
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Password:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1200
         Width           =   1395
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "User Name:"
         Height          =   255
         Left            =   180
         TabIndex        =   14
         Top             =   840
         Width           =   1335
      End
   End
   Begin VB.TextBox txtName 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   5235
   End
   Begin VB.Frame fraActivation 
      Caption         =   "Activation Mode"
      Height          =   1275
      Left            =   0
      TabIndex        =   10
      Top             =   660
      Width           =   5235
      Begin VB.OptionButton optServer 
         Caption         =   "Server Package"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Tag             =   "InProc"
         Top             =   780
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton optLocal 
         Caption         =   "Library Package"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Tag             =   "Local"
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Components will be activated in a dedicated server process."
         Height          =   195
         Left            =   420
         TabIndex        =   13
         Top             =   1020
         Width           =   4275
      End
      Begin VB.Label Label2 
         Caption         =   "Components will be activated in the creator's process."
         Height          =   255
         Left            =   420
         TabIndex        =   12
         Top             =   480
         Width           =   3855
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Enter the name for this package:"
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   2535
   End
End
Attribute VB_Name = "frmAddPackage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_oDOM As DOMDocument30
Private m_oErr As CError

Private Sub cmdAdd_Click()
    Dim lRet As Long
    Dim sXML As String
    Dim oMTSAdmin As WOWMTSAdmin.CMtsAdmin
    Dim oPackage As IXMLDOMNode
    Dim i As Long
    Const PROC_NAME As String = "cmdAdd_Click"
    
    On Error GoTo ErrorHandler
    
    m_oErr.AddProcedureName PROC_NAME
    Me.MousePointer = vbHourglass
    
    'add the properties
    Set oPackage = m_oDOM.selectSingleNode("/Root/Package")
    
    'kill all of the child elements
    While oPackage.hasChildNodes
        oPackage.removeChild oPackage.firstChild
    Wend
    
    If optInteractive = True Then
        AddElement oPackage, "Authentication", "0"
    Else
        'check the password to see if they typed it right twice
        If StrComp(txtPassword.Text, txtConfirm.Text) = 0 Then
            AddElement oPackage, "Password", txtPassword.Text
        Else
            Call MsgBox("Your passwords didn't match.  Please re-type them and try again.", vbOKOnly + vbInformation, "Passwords Error")
            Exit Sub
        End If
        
        AddElement oPackage, "Authentication", "4"
        AddElement oPackage, "Identity", txtUser.Text
        
    End If
    
    AddElement oPackage, "Name", txtName.Text
    AddElement oPackage, "IsSystem", "N"
    
    If optLocal = True Then
        AddElement oPackage, "Activation", "InProc"
    Else
        AddElement oPackage, "Activation", "Local"
    End If
    
    'create an instance of the mts admin on the computer supplied
    If m_oDOM.selectSingleNode("//ServerName").Text = "" Then
        Set oMTSAdmin = CreateObject("WOWMTSAdmin.CMtsAdmin")
    Else
        Set oMTSAdmin = CreateObject("WOWMTSAdmin.CMtsAdmin", m_oDOM.selectSingleNode("//ServerName").Text)
    End If
    
    'get the packages
    lRet = oMTSAdmin.AddPackage(m_oDOM.xml, sXML)
            
    'clean up
    Set oMTSAdmin = Nothing
    Me.MousePointer = vbDefault
    
    If lRet = 0 Then
        Call MsgBox("Package Installed.", vbOKOnly + vbInformation, "Package Installation Status")
    Else
        Call MsgBox("Package was not installed.  Please check to error logs for more information.  The error code was " & lRet & ".", vbCritical + vbOKOnly, "Package Install Status")
    End If
    
    frmMTSAdmin.RefreshTree
    
Exit_Point:
    'return the results
    
    'clean up
    Set oMTSAdmin = Nothing
    Set oPackage = Nothing
        
    'remove us from the error stack
    m_oErr.RemoveProcedureName PROC_NAME
    Unload frmAddPackage
    Exit Sub
ErrorHandler:
    'handle the error
    lRet = Err.Number
    
    m_oErr.HandleError lRet, Err.Description, p_bDisplayMessage:=True
    GoTo Exit_Point
    
End Sub

Public Property Set DOM(ByVal r_oDOM As DOMDocument30)
    Set m_oDOM = r_oDOM
End Property

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

Private Sub Form_Load()
    Set m_oErr = New CError
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set m_oDOM = Nothing
    Set m_oErr = Nothing
End Sub

