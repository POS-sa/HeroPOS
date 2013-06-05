Attribute VB_Name = "modDataStream"
 Private Type COPYDATASTRUCT
              dwData As Long
              cbData As Long
              lpData As Long
      End Type

      Private Const WM_COPYDATA = &H4A

      Public Declare Function FindWindow Lib "user32" Alias _
         "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName _
         As String) As Long

      Private Declare Function SendMessage Lib "user32" Alias _
         "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal _
         wParam As Long, lParam As Any) As Long

      'Copies a block of memory from one location to another.
      Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
         (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

Public Sub send_data_steam_keylog(message As String)
On Error Resume Next
          a$ = DateTime.Now & " (" & UserRecord.User_Number & ") " & UserRecord.Name & message
          Dim cds As COPYDATASTRUCT
          Dim ThWnd As Long
          Dim buf(1 To 255) As Byte

      ' Get the hWnd of the target application
          ThWnd = FindWindow(vbNullString, "Target")
          
      ' Copy the string into a byte array, converting it to ASCII
          Call CopyMemory(buf(1), ByVal a$, Len(a$))
          cds.dwData = 3
          cds.cbData = Len(a$) + 1
          cds.lpData = VarPtr(buf(1))
          i = SendMessage(ThWnd, WM_COPYDATA, frmSplash.hwnd, cds)
End Sub



 Public Function SaveResItemToDisk( _
                ByVal iResourceNum, _
                ByVal sResourceType As String, _
                ByVal sDestFileName As String _
                ) As Long
        '=============================================
        'Saves a resource item to disk

    'Returns 0 on success, error number on failure
    '=============================================

    'Example Call:
    ' iRetVal = SaveResItemToDisk(101, "CUSTOM", "C:\myImage.gif")

    Dim bytResourceData()   As Byte
    Dim iFileNumOut         As Integer

    On Error GoTo SaveResItemToDisk_err

    'Retrieve the resource contents (data) into a byte array
    bytResourceData = LoadResData(iResourceNum, sResourceType)

    'Get Free File Handle
    iFileNumOut = FreeFile

    'Open the output file
    Open sDestFileName For Binary Access Write As #iFileNumOut

        'Write the resource to the file
        Put #iFileNumOut, , bytResourceData

    'Close the file
    Close #iFileNumOut

    'Return 0 for success
    SaveResItemToDisk = 0

    Exit Function
SaveResItemToDisk_err:
    'Return error number
    SaveResItemToDisk = err.Number
End Function

