VERSION 5.00
Begin VB.Form frmGameLoop 
   AutoRedraw      =   -1  'True
   Caption         =   "Game Loop"
   ClientHeight    =   4230
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   282
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   2760
      TabIndex        =   2
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4080
      Top             =   3600
   End
   Begin VB.CommandButton cmdTimer 
      Caption         =   "Start Timer"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton cmdStartLoop 
      Caption         =   "Start Game Loop"
      Height          =   495
      Left            =   1440
      TabIndex        =   0
      Top             =   3600
      Width           =   1215
   End
End
Attribute VB_Name = "frmGameLoop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'APIs
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
'Constants for the GenerateDC function
'**LoadImage Constants**
Const IMAGE_BITMAP As Long = 0
Const LR_LOADFROMFILE As Long = &H10
Const LR_CREATEDIBSECTION As Long = &H2000
'****************************************

Const SpriteWidth As Long = 64
Const SpriteHeight As Long = 64


Dim LastTick As Long
Dim CurrentTick As Long

Dim bLoopRunning As Boolean

Dim Sprite As Long
Dim Mask As Long

Dim X As Long, Y As Long

'IN: FileName: The file name of the graphics
'OUT: The Generated DC
Public Function GenerateDC(FileName As String) As Long
Dim DC As Long
Dim hBitmap As Long

'Create a Device Context, compatible with the screen
DC = CreateCompatibleDC(0)

If DC < 1 Then
    GenerateDC = 0
    'Raise error
    Err.Raise vbObjectError + 1
    Exit Function
End If

'Load the image....BIG NOTE: This function is not supported under NT, there you can not
'specify the LR_LOADFROMFILE flag
hBitmap = LoadImage(0, FileName, IMAGE_BITMAP, 0, 0, LR_LOADFROMFILE Or LR_CREATEDIBSECTION)

If hBitmap = 0 Then 'Failure in loading bitmap
    DeleteDC DC
    GenerateDC = 0
    'Raise error
    Err.Raise vbObjectError + 2
    Exit Function
End If

'Throw the Bitmap into the Device Context
SelectObject DC, hBitmap

'Return the device context
GenerateDC = DC

'Delte the bitmap handle object
DeleteObject hBitmap

End Function
'Deletes a generated DC
Private Function DeleteGeneratedDC(DC As Long) As Long

If DC > 0 Then
    DeleteGeneratedDC = DeleteDC(DC)
Else
    DeleteGeneratedDC = 0
End If

End Function

Private Sub cmdExit_Click()

bLoopRunning = False
CleanUp

End Sub

Private Sub cmdStartLoop_Click()

Timer1.Enabled = False
bLoopRunning = True
RunGameLoop

End Sub

Private Sub cmdTimer_Click()

If bLoopRunning Then bLoopRunning = False

Timer1.Enabled = True

End Sub


Private Sub Form_Load()
On Error GoTo ErrorHandler

Sprite = GenerateDC(App.Path & "\sprite.bmp")
Mask = GenerateDC(App.Path & "\mask.bmp")

ErrorHandler:

Select Case Err

    Case 0 'No error
           
        
    Case Else
        MsgBox "Failed to load graphics"
        CleanUp
End Select

End Sub

Private Sub RunGameLoop()
Const TickDifference As Long = 1

Do
     
    CurrentTick = GetTickCount()
       
    If CurrentTick - LastTick > TickDifference Then
        
        Me.Cls
        'Draw the sprite
        BitBlt Me.hdc, X, Y, SpriteWidth, SpriteHeight, Mask, 0, 0, vbSrcAnd
        BitBlt Me.hdc, X, Y, SpriteWidth, SpriteHeight, Sprite, 0, 0, vbSrcPaint
        
        
        LastTick = GetTickCount()
        
        X = (X Mod Me.ScaleWidth) + 2
        Y = (Y Mod Me.ScaleHeight) + 2
        
        Me.Refresh
        
    Else
        
        'Don't do anything
        
    End If
    
    DoEvents
    
Loop While bLoopRunning

End Sub

Private Sub CleanUp()

DeleteGeneratedDC Sprite
DeleteGeneratedDC Mask

Unload Me
Set frmGameLoop = Nothing

End Sub

Private Sub Timer1_Timer()

Me.Cls

'Draw the sprite
BitBlt Me.hdc, X, Y, SpriteWidth, SpriteHeight, Mask, 0, 0, vbSrcAnd
BitBlt Me.hdc, X, Y, SpriteWidth, SpriteHeight, Sprite, 0, 0, vbSrcPaint
        
X = (X Mod Me.ScaleWidth) + 2
Y = (Y Mod Me.ScaleHeight) + 2

Me.Refresh

End Sub
