Attribute VB_Name = "basLocalUNC"
Option Explicit

'Get a UNC path for a Shared Local Directory
'***Much of this code was derived from Karl E. Peterson  http://www.mvps.org/vb***
'****With some modifications.  Check out His Sight its Really Helpful !***
'Constants
Private Const ERROR_ACCESS_DENIED = 5&
Private Const LM20_NNLEN = 12         ' // LM 2.0 Net name length
Private Const MAX_PREFERRED_LENGTH As Long = -1
Private Const NO_ERROR = 0
Private Const SHPWLEN = 8             ' // Share password length (bytes)
Private Const STYPE_DISKTREE = 0      ' /* disk share */
Private Const VER_PLATFORM_WIN32_NT = 2


Private mbUNCNeedAdminPrivs As Boolean

' Need UDTs for Checking UNC path of Local Directory

Private Type REMOTE_NAME_INFO
   lpUniversalName As Long
   lpConnectionName As Long
   lpRemainingPath As Long
End Type

Private Type SHARE_INFO_2
   Netname As String
   ShareType As Long
   Remark As String
   Permissions As Long
   MaxUsers As Long
   CurrentUsers As Long
   Path As String
   Password As String
End Type

Private Type SHARE_INFO_50            'struct _share_info_50 {
   Netname(0 To LM20_NNLEN) As Byte   '   char            shi50_netname[LM20_NNLEN+1];
   ShareType As Byte                  '   unsigned char   shi50_type;
   Flags As Integer                   '   unsigned short  shi50_flags;
   lpRemark As Long                   '   char FAR *      shi50_remark;
   lpPath As Long                     '   char FAR *      shi50_path;
   PasswordRW(0 To SHPWLEN) As Byte   '   char            shi50_rw_password[SHPWLEN+1];
   PasswordRO(0 To SHPWLEN) As Byte   '   char            shi50_ro_password[SHPWLEN+1];
End Type

Private Type OSVERSIONINFORMATION
   dwOSVersionInfoSize As Long
   dwMajorVersion As Long
   dwMinorVersion As Long
   dwBuildNumber As Long
   dwPlatformId As Long
   szCSDVersion As String * 128
End Type

' Need these APIs, part of proccess to get UNC for Local Directory
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pTo As Any, uFrom As Any, ByVal lSize As Long)
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFORMATION) As Long
Private Declare Function lstrlenA Lib "kernel32" (ByVal PointerToString As Long) As Long
Private Declare Function lstrlenW Lib "kernel32" (ByVal PointerToString As Long) As Long
Private Declare Function NetApiBufferFree Lib "netapi32" (ByVal lpBuffer As Long) As Long
Private Declare Function NetShareEnum Lib "netapi32" (ByVal lpServerName As Long, ByVal dwLevel As Long, lpBuffer As Any, ByVal dwPrefMaxLen As Long, EntriesRead As Long, TotalEntries As Long, hResume As Long) As Long
Private Declare Function NetShareEnum95 Lib "svrapi" Alias "NetShareEnum" (ByVal lpServerName As String, ByVal dwLevel As Long, lpBuffer As Any, ByVal cbBuffer As Long, EntriesRead As Long, TotalEntries As Long) As Long
'END



'Following Functions are used to get UNC for Shared Local Directory
'Public Function GetLocalUncName
'Private Function CurrentMachineName
'Private Function PointerToDWord
'Private Function PointerToStringW
'Public Function EnumShares
'Public Function IsWinNT
'Private Function EnumSharesNT
'Private Function EnumShares9x
'Private Function TrimNull
'Private Function PointerToStringA





'This Function will return the UNC path of a Shared "Local Directory ONLY", Local Path if directory is not shared.
'This function would be good to use in a select case of another function to Get UNC path in Case you get
'the Error WN_NOT_CONNECTED ErrorMsg = "The drive is not connected"  In other words you are attempting to get
'the UNC path of a "Local directory" not a Mapped directory :)
Public Function GetLocalUncName(ByVal sFileSpec As String) As String
    Dim Buffer() As Byte
    Dim lRet As Long
    Dim shi() As SHARE_INFO_2
    Dim lCount As Long
    
    GetLocalUncName = sFileSpec
    ' get list of local shares
    lRet = EnumShares(shi)
    If lRet > 0 Then
       ' loop through shares, looking for a potential match
       ' ambiguous: any path can be on more than one share
        For lCount = 0 To lRet - 1
            If shi(lCount).ShareType = STYPE_DISKTREE Then
                If InStr(1, sFileSpec, shi(lCount).Path, vbTextCompare) = 1 Then
                    ' this element starts with the same path
                    ' have to accept first match
'                    GetLocalUncName = "\\" & CurrentMachineName & _
'                                    "\" & shi(lCount).Netname & _
'                                    Mid(sFileSpec, Len(shi(lCount).Path) + 1)
                    'Dont need the +1 in Mid(sFileSpec, Len(shi(lCount).Path)+1)
                    'if The Local Drive is Shared.  If Only a folder in the Local Drive is Shared,
                    'and the Drive itself is not shared, then you need the +1 :)
                    GetLocalUncName = "\\" & CurrentMachineName & "\" & shi(lCount).Netname
                    If UCase(Left(sFileSpec, Len(sFileSpec) - 1)) = UCase(shi(lCount).Path) Then
                        GetLocalUncName = GetLocalUncName & Mid(sFileSpec, Len(shi(lCount).Path) + 1)
                    Else
                        GetLocalUncName = GetLocalUncName & Mid(sFileSpec, Len(shi(lCount).Path))
                    End If
                    Exit For
                End If
            End If
        Next lCount
    End If
End Function

' Continued ... This Function uses GetComputerName API
Private Function CurrentMachineName() As String
   Dim Buffer As String
   Dim nLen As Long
   Const CNLEN = 15          ' Maximum computer name length
   
   Buffer = Space(CNLEN + 1)
   nLen = Len(Buffer)
   If GetComputerName(Buffer, nLen) Then
      CurrentMachineName = Left(Buffer, nLen)
   End If
End Function

' Continued ... Uses CopyMemory API
Private Function PointerToDWord(ByVal lpDWord As Long) As Long
   Call CopyMemory(PointerToDWord, ByVal lpDWord, 4)
End Function

'Continued ... Uses lstrlenW API and CopyMemory API
Private Function PointerToStringW(ByVal lpStringW As Long) As String
   Dim Buffer() As Byte
   Dim nLen As Long
   
   If lpStringW Then
      nLen = lstrlenW(lpStringW) * 2
      If nLen Then
         ReDim Buffer(0 To (nLen - 1)) As Byte
         CopyMemory Buffer(0), ByVal lpStringW, nLen
         PointerToStringW = Buffer
      End If
   End If
End Function

'Continued ... This Functions checks to see If the current platform
'is NT then decides which EnumShares to use to get an array of current SHared directories
'are set up on the local machine
Public Function EnumShares(shi() As SHARE_INFO_2, Optional ByVal Server As String = "") As Long
   Dim nRet As Long
   ' delegate to OS-appropriate routine
   If IsWinNT() Then
      nRet = EnumSharesNT(shi(), Server)
      mbUNCNeedAdminPrivs = (nRet < 0)
   Else
      nRet = EnumShares9x(shi(), Server)
   End If
   EnumShares = nRet
End Function

'Continued ... This function Returns true if the Current platform is
'WInNT using the Existing API call located in WINAPI.BAS It is called by EnumShares
Public Function IsWinNT() As Boolean
   Static os As OSVERSIONINFORMATION
   Static bRet As Boolean
   ' just do this once, for optimization
   If os.dwPlatformId = 0 Then
      os.dwOSVersionInfoSize = Len(os)
      Call GetVersionEx(os)
      bRet = (os.dwPlatformId = VER_PLATFORM_WIN32_NT)
   End If
   IsWinNT = bRet
End Function
' Continued ... This function Fills the shi array with UTD SHARE_INFO_2
'So for every instance of a shared (folder) directory on a WINNT Platform it will Populate all
'the data concerning SHARE_INFO_2
Private Function EnumSharesNT(shi() As SHARE_INFO_2, Optional ByVal Server As String = "") As Long
   Dim Level As Long
   Dim lpBuffer As Long
   Dim EntriesRead As Long
   Dim TotalEntries As Long
   Dim hResume As Long
   Dim Offset As Long
   Dim nRet As Long
   Dim i As Long
   
   ' convert Server to null pointer if none requested.
   ' this has the effect of asking for the local machine.
   If Len(Server) = 0 Then Server = vbNullString

   ' ask for all available shares; try level 2 first
   Level = 2
   nRet = NetShareEnum(StrPtr(Server), Level, lpBuffer, MAX_PREFERRED_LENGTH, EntriesRead, TotalEntries, hResume)
   
   If nRet = ERROR_ACCESS_DENIED Then
      ' bummer -- need admin privs for level 2, drop to level 1
      Level = 1
      nRet = NetShareEnum(StrPtr(Server), Level, lpBuffer, MAX_PREFERRED_LENGTH, EntriesRead, TotalEntries, hResume)
   End If
   
   If nRet = NO_ERROR Then
      ' make sure there are shares to decipher
      If EntriesRead > 0 Then
         ' prepare UDT buffer to hold all share info
         ReDim shi(0 To EntriesRead - 1)
         ' loop through API buffer, extracting each element
         For i = 0 To EntriesRead - 1
            With shi(i)
               .Netname = PointerToStringW(PointerToDWord(lpBuffer + Offset))
               .ShareType = PointerToDWord(lpBuffer + Offset + 4)
               .Remark = PointerToStringW(PointerToDWord(lpBuffer + Offset + 8))
               If Level = 2 Then
                  .Permissions = PointerToDWord(lpBuffer + Offset + 12)
                  .MaxUsers = PointerToDWord(lpBuffer + Offset + 16)
                  .CurrentUsers = PointerToDWord(lpBuffer + Offset + 20)
                  .Path = PointerToStringW(PointerToDWord(lpBuffer + Offset + 24))
                  .Password = PointerToStringW(PointerToDWord(lpBuffer + Offset + 28))
                  Offset = Offset + Len(shi(i))
               Else
                  Offset = Offset + 12  ' Len(SHARE_INFO_1)
               End If
            End With
         Next i
      End If
      
      ' return number of entries found
      If Level = 1 Then
         ' negative if we don't have admin privs
         EnumSharesNT = -EntriesRead
      ElseIf Level = 2 Then
         EnumSharesNT = EntriesRead
      End If
   End If
   
   ' clean up
   If lpBuffer Then
      Call NetApiBufferFree(lpBuffer)
   End If
End Function

'Continued ... This function Fills the shi array with UTD SHARE_INFO_2
'So for every instance of a shared (folder) directory on a WIN9x Platform it will Populate all
'the data concerning SHARE_INFO_2
Private Function EnumShares9x(shi() As SHARE_INFO_2, Optional ByVal Server As String = "") As Long
   Dim Buffer() As Byte
   Dim EntriesRead As Long
   Dim TotalEntries As Long
   Dim Offset As Long
   Dim shi95 As SHARE_INFO_50
   Dim nRet As Long
   Dim i As Long
   Const BufferSize = &H4000
   
   ' convert Server to null pointer if none requested.
   ' this has the effect of asking for the local machine.
   If Len(Server) = 0 Then Server = vbNullString

   ' ask for all available shares, using really large buffer
   ReDim Buffer(0 To BufferSize - 1) As Byte
   nRet = NetShareEnum95(Server, 50, Buffer(0), BufferSize, EntriesRead, TotalEntries)
   
   If nRet = NO_ERROR Then
      ' make sure there are shares to decipher
      If EntriesRead > 0 Then
         ' prepare UDT buffer to hold all share info
         ReDim shi(0 To EntriesRead - 1)
         ' loop through API buffer, extracting each element
         For i = 0 To EntriesRead - 1
            With shi(i)
               Call CopyMemory(shi95, Buffer(Offset), Len(shi95))
               .Netname = TrimNull(StrConv(shi95.Netname, vbUnicode))
               .ShareType = shi95.ShareType
               .Path = PointerToStringA(shi95.lpPath)
               .Remark = PointerToStringA(shi95.lpRemark)
               If shi95.PasswordRW(0) = 0 Then
                  .Password = TrimNull(StrConv(shi95.PasswordRO, vbUnicode))
               Else
                  .Password = TrimNull(StrConv(shi95.PasswordRW, vbUnicode))
               End If
               
               Offset = Offset + Len(shi95)
            End With
         Next i
         
         ' return number of entries found
         EnumShares9x = EntriesRead
      End If
   End If
End Function

'Continued Need this function to Truncate String Checks for Null first

Private Function TrimNull(ByVal StrIn As String) As String
   Dim nul As Long
   '
   ' Truncate input string at first null.
   ' If no nulls, perform ordinary Trim.
   '
   nul = InStr(StrIn, vbNullChar)
   Select Case nul
      Case Is > 1
         TrimNull = Left(StrIn, nul - 1)
      Case 1
         TrimNull = ""
      Case 0
         TrimNull = Trim(StrIn)
   End Select
End Function

'Continued ... Uses lstrlenA API and CopyMemory API

Private Function PointerToStringA(lpStringA As Long) As String
   Dim Buffer() As Byte
   Dim nLen As Long
   
   If lpStringA Then
      nLen = lstrlenA(ByVal lpStringA)
      If nLen Then
         ReDim Buffer(0 To (nLen - 1)) As Byte
         CopyMemory Buffer(0), ByVal lpStringA, nLen
         PointerToStringA = StrConv(Buffer, vbUnicode)
      End If
   End If
End Function
'END Functions to Get UNC for a local directory




