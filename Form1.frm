'Copyright (C) 2009 by Rizky Hafitsyah'

VERSION 5.00
Begin VB.Form FormMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Baca Isi File Text"
   ClientHeight    =   7425
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7425
   ScaleWidth      =   13500
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Baca Dari File"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10920
      TabIndex        =   1
      Top             =   6600
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6135
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Form1.frx":0000
      Top             =   120
      Width           =   13215
   End
End
Attribute VB_Name = "FormMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
On Error GoTo HancurkanError
    Dim i As Integer
    Dim s As String, s1 As String
        i = FreeFile
            Open App.Path & "\test.txt" For Input As #i
                Do Until EOF(i)
                    Input #i, s
                    s1 = s1 & s & IIf(EOF(i), "", vbCrLf)
                Loop
            Close #i
    With Text1
        .Text = s1
        .SetFocus
    End With
Exit Sub

HancurkanError:
    If Err.Number = 53 Then
        MsgBox "Maaf, file tidak ditemukan!", vbExclamation + vbOKOnly, "Error"
        With Text1
            .SetFocus
        End With
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub Form_Load()
    With Text1
        .Text = Empty
    End With
End Sub
