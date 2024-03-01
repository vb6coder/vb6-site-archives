Attribute VB_Name = "modHelp16_32"
#If Win16 Then
    
  Declare Function WinHelp Lib "user" _
    (ByVal hWnd As Integer, _
    ByVal lpHelpFile As String, _
    ByVal wCommand As Integer, _
    ByVal dwData As Long) As Integer
  
  Declare Function WinHelpTopic Lib "user" Alias "WinHelp" _
    (ByVal hWnd As Integer, _
    ByVal lpHelpFile As String, _
    ByVal wCommand As Integer, _
    ByVal dwData As String) As Integer
    
#Else
  Declare Function WinHelp Lib "user32" Alias "WinHelpA" _
    (ByVal hWnd As Long, _
    ByVal lpHelpFile As String, _
    ByVal wCommand As Long, _
    ByVal dwData As Long) As Long
  
  Declare Function WinHelpTopic Lib "user32" Alias "WinHelpA" _
    (ByVal hWnd As Long, _
    ByVal lpHelpFile As String, _
    ByVal wCommand As Long, _
    ByVal dwData As String) As Long
#End If

Public Sub ShowHelpContents(ByVal intHelpFile As Integer)
  
  #If Win16 Then
    ' Open the help file using the
    ' 16-bit HELP_CONTENTS constant (3)
    WinHelp hWnd, SetHelpStrings(intHelpFile), 3, 0
  #Else
  
    ' Open the help file using the
    ' 32-bit HELP_TAB constant (15)
    WinHelp hWnd, SetHelpStrings(intHelpFile), 15, 0
  #End If
  
End Sub

Public Sub ShowHelpIndex(ByVal intHelpFile As Integer)
  
  ' Open the help file using the
  ' HELP_PARTIALKEY constant (261)
  WinHelpTopic hWnd, SetHelpStrings(intHelpFile), 261, ""
    
End Sub

Public Sub ShowHelpTopic(ByVal intHelpFile As Integer, strTopic As String)
  
  ' Open the help file using the
  ' HELP_KEY constant (&H101)
  WinHelpTopic hWnd, SetHelpStrings(intHelpFile), &H101, strTopic
  
End Sub

Public Sub ShowHelpContextID(ByVal intHelpFile As Integer, lngContextID As Long)
  
  ' Open the help file using the
  ' HELP_CONTEXT constant (1)
  WinHelp hWnd, SetHelpStrings(intHelpFile), 1, lngContextID
  
End Sub

Public Sub ShowHelpSecondary(ByVal intHelpFile As Integer, lngContextID As Long)
  
  ' Open the help file in a secondary
  ' window named "subMain" using the
  ' HELP_CONTEXT constant (1)
  strHelpWindow = SetHelpStrings(intHelpFile) & ">subMain"
  WinHelp hWnd, strHelpWindow, 1, lngContextID

End Sub

Public Sub ShowHelpKeyword(ByVal intHelpFile As Integer, strKeyword As String)
  
  ' Open the help file using the
  ' HELP_PARTIALKEY constant (&H105)
  WinHelpTopic hWnd, SetHelpStrings(intHelpFile), &H105, strKeyword

End Sub


Public Function SetHelpStrings(ByVal intSelHelpFile As Integer) As String
  
  ' Set the string variable to
  ' include the application path
  Select Case intSelHelpFile
  Case 1
    SetHelpStrings = App.Path & "\Help\Barcode.hlp"
  Case 2
    ' Place other Help file paths in other Case statements
  End Select
  
End Function

Public Sub ShowHelpPartialKey(ByVal intHelpFile As Integer, strPartKey As String)

  ' Open the help file using the
  ' HELP_PARTIALKEY constant (261)
  WinHelpTopic hWnd, SetHelpStrings(intHelpFile), 261, strPartKey
  
End Sub