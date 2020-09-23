VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5850
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7740
   LinkTopic       =   "Form1"
   ScaleHeight     =   5850
   ScaleWidth      =   7740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdEnd 
      Caption         =   "&End"
      Height          =   495
      Left            =   6360
      TabIndex        =   0
      Top             =   5040
      Width           =   1095
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function BitBlt Lib "gdi32" _
                                    (ByVal hDestDC As Long, _
                                     ByVal x As Long, _
                                     ByVal y As Long, _
                                     ByVal nWidth As Long, _
                                     ByVal nHeight As Long, _
                                     ByVal hSrcDC As Long, _
                                     ByVal xSrc As Long, _
                                     ByVal ySrc As Long, _
                                     ByVal dwRop As Long) As Long


Private Sub cmdEnd_Click()
End
End Sub

Private Sub Form_Load()
Dim A As Long
Dim s As Long
frmMain.AutoRedraw = True
frmMain.ScaleMode = 1
A = GetDesktopWindow()
s = GetDC(A)
BitBlt Me.hDC, 0, 0, Screen.Width, Screen.Height, s, 0, 0, vbSrcCopy
End Sub


