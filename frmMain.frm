VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "128bit Barcode"
   ClientHeight    =   1530
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6630
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   6630
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picBarcode 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   4680
      ScaleHeight     =   12
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   11
      TabIndex        =   4
      Top             =   660
      Width           =   195
   End
   Begin VB.PictureBox picBarcodeLarge 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   1080
      Left            =   5400
      ScaleHeight     =   70
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   2
      Top             =   240
      Width           =   990
   End
   Begin VB.TextBox txtInput 
      Height          =   1080
      Left            =   720
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   240
      Width           =   2880
   End
   Begin VB.Label Label1 
      Caption         =   "Result:             -->"
      Height          =   255
      Index           =   1
      Left            =   3960
      TabIndex        =   3
      Top             =   660
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Text Input:"
      Height          =   495
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   540
      Width           =   615
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type RECT
    Left   As Long
    Top    As Long
    Right  As Long
    Bottom As Long
End Type

Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function FillRect Lib "user32.dll" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function SetPixel Lib "gdi32.dll" Alias "SetPixelV" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long

Dim Bits(7) As Long

Private Sub Encode()
    Dim cMD5              As clsMD5
    Dim sMD5Hash          As String
    Dim blBitArray(127)   As Boolean
    Dim a, b, iX, iY      As Integer
    Dim tRect             As RECT
    Dim lBrush            As Long
    Dim iSize             As Integer
    
    If txtInput.Text <> "" Then
        Set cMD5 = New clsMD5
        sMD5Hash = cMD5.MD5(txtInput.Text)
        Set cMD5 = Nothing
        
        For a = 0 To 15
            For b = 0 To 7
                blBitArray((a * 8) + b) = CByte("&H" & Mid$(sMD5Hash, ((a + 1) * 2), 2)) And Bits(b)
            Next b
        Next a
        
        iX = 0: iY = 0
        iSize = 6
        
        picBarcode.Cls: picBarcodeLarge.Cls
        
        lBrush = CreateSolidBrush(&H0&)
        For iY = 0 To 11
            For iX = 1 To 11
                If blBitArray((iY * 10) + iX) Then Call SetPixel(picBarcode.hDC, iX - 1, iY, &H0&)
                
                With tRect
                    .Left = (iX * iSize) - iSize
                    .Top = (iY * iSize)
                    .Right = (.Left + iSize)
                    .Bottom = (.Top + iSize)
                End With
                If blBitArray((iY * 10) + iX) Then Call FillRect(picBarcodeLarge.hDC, tRect, lBrush)
            Next iX
        Next iY
        Call DeleteObject(lBrush)
    Else
        picBarcode.Cls: picBarcodeLarge.Cls
    End If
End Sub

Private Sub Form_Load()
    Dim I As Integer
    
    For I = 0 To 7
        Bits(I) = (2 ^ I): DoEvents
    Next I
End Sub

Private Sub txtInput_Change()
    Call Encode
End Sub
