VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Memeriksa Apakah Printer Terinstall di PC Anda"
   ClientHeight    =   3090
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   7050
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   7050
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Periksa!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1560
      TabIndex        =   0
      Top             =   960
      Width           =   3735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public Function IsPrinterInstalled() As Boolean
    On Error Resume Next
    
    Dim strDummy As String
    Dim PrinterInstalled As Boolean
    
    strDummy = Printer.DeviceName
      If Err.Number Then
         PrinterInstalled = False
      Else
         PrinterInstalled = True
      End If
End Function

Private Sub Command1_Click()
    If IsPrinterInstalled Then
       MsgBox "Printer terinstall di PC Anda!", _
               vbInformation, "Terinstall"
    Else
       MsgBox "Printer belum terinstall di PC Anda!", _
               vbCritical, "Belum Terinstall"
    End If
End Sub



