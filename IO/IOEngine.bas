Attribute VB_Name = "IO"
' ===========================================
' IO System (Input/Output) [Keyboard & Mouse]
' ===========================================
' RokyBeast@RokyXxX (MIT License)
' Take KeyCode References from: https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/keycode-constants

Option Explicit

#If VBA7 Then
    Public Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
#Else
    Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
#End If

Public vRunning As Boolean
Public Const delayMs As Integer = 10

Public Property Let IOState(NewIO As Boolean)
    vRunning = NewIO
End Property

Public Sub InitIO()
    IOState = True
    Call IOHandler
End Sub

Public Sub EndIO()
    IOState = False
End Sub

' Main Sub
Public Sub IOHandler()
    While vRunning
        If GetAsyncKeyState(KEY_NAME) And &H8000 Then
            ' Code
        End If
        
        DoEvents
        Delay delayMs
    Wend
End Sub

Private Sub Delay(Optional ms As Long = 10)
    Dim t As Single: t = Timer
    Do While Timer - t < ms / 1000
        DoEvents
    Loop
End Sub
