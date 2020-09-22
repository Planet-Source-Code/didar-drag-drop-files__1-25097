Attribute VB_Name = "dragdrop"

   Type POINTAPI
      x As Long
      y As Long
   End Type


   Type MSG
      hWnd As Long
      message As Long
      wParam As Long
      lParam As Long
      time As Long
      pt As POINTAPI
   End Type

   Public Declare Sub DragAcceptFiles Lib "shell32.dll" _
     (ByVal hWnd As Long, ByVal fAccept As Long)

   Public Declare Sub DragFinish Lib "shell32.dll" _
     (ByVal hDrop As Long)

   Public Declare Function DragQueryFile Lib "shell32.dll" Alias "DragQueryFileA" _
     (ByVal hDrop As Long, ByVal UINT As Long, _
      ByVal lpStr As String, ByVal ch As Long) As Long

   Public Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" _
     (lpMsg As MSG, ByVal hWnd As Long, ByVal wMsgFilterMin As Long, _
      ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long

   Public Const PM_NOREMOVE = &H0
   Public Const PM_NOYIELD = &H2
   Public Const PM_REMOVE = &H1
   Public Const WM_DROPFILES = &H233





Public Sub Main()
'drag.Show
Form2.Show
WatchForFiles

End Sub

Public Sub WatchForFiles()
   Dim FileDropMessage As MSG
   Dim fileDropped As Boolean
   Dim hDrop As Long
   Dim filename As String * 128
   Dim numOfDroppedFiles As Long
   Dim curFile As Long
 
   Do
       fileDropped = PeekMessage(FileDropMessage, 0, _
                     WM_DROPFILES, WM_DROPFILES, PM_REMOVE Or PM_NOYIELD)

       If fileDropped Then
          hDrop = FileDropMessage.wParam
          numOfDroppedFiles = DragQueryFile(hDrop, True, filename, 127)

          For curFile = 1 To numOfDroppedFiles
              ret = DragQueryFile(hDrop, curFile - 1, filename, 127)
              drag.lblNumDropped = LTrim$(Str$(numOfDroppedFiles))
             drag.List1.AddItem filename
          Next curFile
          DragFinish (hDrop)

   End If
   DoEvents

   Loop

End Sub



