VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAddComponent 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Install New Components"
   ClientHeight    =   1545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6090
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1545
   ScaleWidth      =   6090
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdInstall 
      Caption         =   "Install"
      Height          =   375
      Left            =   4980
      TabIndex        =   3
      Top             =   1020
      Width           =   1035
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
      Height          =   375
      Left            =   4980
      TabIndex        =   2
      Top             =   480
      Width           =   1035
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add..."
      Height          =   375
      Left            =   4980
      TabIndex        =   1
      Top             =   60
      Width           =   1035
   End
   Begin VB.ListBox lstComponents 
      Height          =   1425
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4695
   End
   Begin MSComDlg.CommonDialog dlg1 
      Left            =   5160
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Select files to install"
      Filter          =   "Component Files (*.tlb;*.dll)|*.tlb;*.dll|All Files (*.*)|*.*"
   End
End
Attribute VB_Name = "frmAddComponent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_oDOM As DOMDocument30
Private m_oErr As CError

Private Sub cmdAdd_Click()
    Dim sRet As String

    Const CANCELERROR As Long = 32755
      
    On Error Resume Next
    
    dlg1.ShowOpen
    
    If Not Err.Number = CANCELERROR Then
        sRet = GetUncName(dlg1.FileName)
        If StrComp(sRet, dlg1.FileName, vbTextCompare) = 0 Then
            lstComponents.AddItem dlg1.FileName
        Else
            lstComponents.AddItem sRet
        End If
    End If
End Sub

Public Property Set DOM(ByVal r_oDOM As DOMDocument30)
    Set m_oDOM = r_oDOM
End Property

Private Sub cmdInstall_Click()
    Dim lRet As Long
    Dim sXML As String
    Dim oMTSAdmin As WOWMTSAdmin.CMtsAdmin
    Dim oPackage As IXMLDOMNode
    Dim oNode As IXMLDOMNode
    Dim i As Long
    Const PROC_NAME As String = "cmdInstall_Click"
    On Error GoTo ErrorHandler
    m_oErr.AddProcedureName PROC_NAME
    Me.MousePointer = vbHourglass
    
    'add the filename
    Set oPackage = m_oDOM.selectSingleNode("/Root/Package")
    For i = 0 To lstComponents.ListCount - 1
        AddElement oPackage, "Component", "", oNode
        AddElement oNode, "DLL", lstComponents.List(i)
    Next i
        
    'create an instance of the mts admin on the computer supplied
    If m_oDOM.selectSingleNode("//ServerName").Text = "" Then
        Set oMTSAdmin = CreateObject("WOWMTSAdmin.CMtsAdmin")
    Else
        Set oMTSAdmin = CreateObject("WOWMTSAdmin.CMtsAdmin", m_oDOM.selectSingleNode("//ServerName").Text)
    End If
    
    'get the packages
    lRet = oMTSAdmin.AddDLL(m_oDOM.xml, sXML)
            
    'clean up
    Set oMTSAdmin = Nothing
    Me.MousePointer = vbDefault
    
    If lRet = 0 Then
        Call MsgBox("Dll Installed.", vbOKOnly + vbInformation, "Dll Installation Status")
    Else
        Call MsgBox("DLL was not installed.  Please check to error logs for more information.  The error code was " & lRet & ".", vbCritical + vbOKOnly, "DLL Install Status")
    End If
    
    frmMTSAdmin.RefreshTree
        
Exit_Point:
    
    'clean up
    Set oMTSAdmin = Nothing
    Set oPackage = Nothing
    Set oNode = Nothing
    
    'remove us from the error stack
    m_oErr.RemoveProcedureName PROC_NAME
    Unload frmAddComponent
    Exit Sub
ErrorHandler:
    'handle the error
    lRet = Err.Number
    
    m_oErr.HandleError lRet, Err.Description, p_bDisplayMessage:=True
    GoTo Exit_Point

End Sub
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

Private Sub cmdRemove_Click()
    If lstComponents.ListIndex <> -1 Then
        lstComponents.RemoveItem lstComponents.ListIndex
    End If
End Sub

Private Sub Form_Load()
    Set m_oErr = New CError
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set m_oDOM = Nothing
    Set m_oErr = Nothing
End Sub
