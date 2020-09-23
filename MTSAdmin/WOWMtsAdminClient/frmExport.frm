VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmExport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Export Package"
   ClientHeight    =   1545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6720
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1545
   ScaleWidth      =   6720
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExport 
      Caption         =   "Export"
      Height          =   315
      Left            =   5460
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog dlg1 
      Left            =   6000
      Top             =   60
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "*.pak"
      DialogTitle     =   "Export Package to File"
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse..."
      Height          =   315
      Left            =   5460
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox txtFileName 
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Top             =   840
      Width           =   5295
   End
   Begin VB.Label Label1 
      Caption         =   $"frmExport.frx":0000
      Height          =   615
      Left            =   60
      TabIndex        =   3
      Top             =   180
      Width           =   6615
   End
End
Attribute VB_Name = "frmExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_oDOM As DOMDocument30
Private m_oErr As CError
Private Sub cmdBrowse_Click()
    Const CANCELERROR As Long = 32755
    
    dlg1.Filter = "Package Files--NT 4.0 (*.pak)|*.pak|Application Files--Win2k (*.msi)|*.msi"
    
    On Error Resume Next
    
    dlg1.ShowOpen
    
    On Error GoTo 0
    If Not Err.Number = CANCELERROR Then
        txtFileName.Text = dlg1.FileName
    End If
End Sub

Public Property Set DOM(ByVal r_oDOM As DOMDocument30)
    Set m_oDOM = r_oDOM
End Property

Private Sub cmdExport_Click()
    Dim lRet As Long
    Dim sXML As String
    Dim oMTSAdmin As WOWMTSAdmin.CMtsAdmin
    Const PROC_NAME As String = "cmdExport_Click"
    
    On Error GoTo ErrorHandler
    
    m_oErr.AddProcedureName PROC_NAME
    
    Me.MousePointer = vbHourglass
    
    'add the filename
    AddElement m_oDOM.selectSingleNode("/Root/Package"), "FileName", txtFileName.Text
        
    'create an instance of the mts admin on the computer supplied
    If m_oDOM.selectSingleNode("//ServerName").Text = "" Then
        Set oMTSAdmin = CreateObject("WOWMTSAdmin.CMtsAdmin")
    Else
        Set oMTSAdmin = CreateObject("WOWMTSAdmin.CMtsAdmin", m_oDOM.selectSingleNode("//ServerName").Text)
    End If
    
    'get the packages
    lRet = oMTSAdmin.ExportPackage(m_oDOM.xml, sXML)
                
Exit_Point:
    'return the results
    Me.MousePointer = vbDefault
    
    If lRet = 0 Then
        Call MsgBox("Package exported", vbOKOnly + vbInformation, "Package Export Status")
    Else
        Call MsgBox("Package was not exported.  Please check to error logs for more information.  The error code was " & lRet & ".", vbCritical + vbOKOnly, "Package Export Status")
    End If
    
    'clean up
    Set oMTSAdmin = Nothing
       
    'remove us from the error stack
    m_oErr.RemoveProcedureName PROC_NAME
    Unload Me
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

Private Sub Form_Load()
    Set m_oErr = New CError
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set m_oDOM = Nothing
    Set m_oErr = Nothing
End Sub
