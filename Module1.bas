Attribute VB_Name = "Module1"
Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOPENFILENAME As _
       OPENFILENAME) As Long


Type OPENFILENAME
  lStructSize As Long
  hwndOwner As Long
  hInstance As Long
  lpstrFilter As String
  lpstrCustomFilter As String
  nMaxCustomFilter As Long
  nFilterIndex As Long
  lpstrFile As String
  nMaxFile As String
  lpstrFileTitle As String
  nMaxFileTitle As String
  lpstrInitialDir As String
  lpstrTitle As String
  flags As Long
  nFileOffset As Integer
  nFileExtension As Integer
  lpstrDefExt As String
  lCustData As Long
  lpfnHook As Long
  lpTemplateName As String
End Type


