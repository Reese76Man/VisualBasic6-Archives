VERSION 5.00
Begin VB.Form frmOldDB 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Old DB Access"
   ClientHeight    =   9660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9660
   ScaleWidth      =   13590
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtSearch 
      Height          =   285
      Left            =   1440
      TabIndex        =   26
      Top             =   120
      Width           =   2535
   End
   Begin VB.HScrollBar RecScroll 
      Height          =   615
      LargeChange     =   25
      Left            =   1440
      Min             =   1
      TabIndex        =   13
      Top             =   5040
      Value           =   1
      Width           =   12015
   End
   Begin VB.TextBox txtRecordText 
      Height          =   285
      Index           =   12
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   4560
      Width           =   12000
   End
   Begin VB.TextBox txtRecordText 
      Height          =   285
      Index           =   11
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   4200
      Width           =   12000
   End
   Begin VB.TextBox txtRecordText 
      Height          =   285
      Index           =   10
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   3840
      Width           =   12000
   End
   Begin VB.TextBox txtRecordText 
      Height          =   285
      Index           =   9
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   3480
      Width           =   12000
   End
   Begin VB.TextBox txtRecordText 
      Height          =   285
      Index           =   8
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   3120
      Width           =   12000
   End
   Begin VB.TextBox txtRecordText 
      Height          =   285
      Index           =   7
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   2760
      Width           =   12000
   End
   Begin VB.TextBox txtRecordText 
      Height          =   285
      Index           =   6
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   2400
      Width           =   12000
   End
   Begin VB.TextBox txtRecordText 
      Height          =   285
      Index           =   5
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   2040
      Width           =   12000
   End
   Begin VB.TextBox txtRecordText 
      Height          =   285
      Index           =   4
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1680
      Width           =   12000
   End
   Begin VB.TextBox txtRecordText 
      Height          =   285
      Index           =   3
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1320
      Width           =   12000
   End
   Begin VB.TextBox txtRecordText 
      Height          =   285
      Index           =   2
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   960
      Width           =   12000
   End
   Begin VB.TextBox txtRecordText 
      Height          =   285
      Index           =   1
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   600
      Width           =   12000
   End
   Begin VB.Label Label2 
      Caption         =   "Search:"
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   120
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      Height          =   5295
      Left            =   0
      Top             =   480
      Width           =   13575
   End
   Begin VB.Label Label1 
      Caption         =   "Phone Number:"
      Height          =   255
      Index           =   11
      Left            =   120
      TabIndex        =   24
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Address:"
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   23
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Misc3:"
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   22
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Misc2:"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   21
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Misc1:"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   20
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Number:"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   19
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Grade:"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   18
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "BirthDate:"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   17
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Created:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   16
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "MiddleName:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   15
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "LastName:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   14
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "FirstName:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   12
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "frmOldDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DataBase(5000, 12) As String, RNum As Long

Private Sub Form_Load()
LoadData

LoadRecord 1
End Sub

Private Sub txtSearch_Change()
If RecScroll.Value > 1 Then LoadRecord 1
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    If Len(Trim(txtSearch.Text)) > 3 Then SearchRecord LCase(Trim(txtSearch.Text))
End If
End Sub

Private Sub RecScroll_Change()
LoadRecord RecScroll.Value
End Sub

Public Sub LoadRecord(ByVal RecordNum As Long)
Dim Cnt As Long
For Cnt = 1 To 12
    txtRecordText(Cnt).Text = Trim(DataBase(RecordNum, Cnt))
Next Cnt
Me.Caption = "Record: " & Trim(Str(RecordNum)) & " of " & Trim(Str(RNum)) & " Total."
RecScroll.Value = RecordNum
End Sub

Public Sub SearchRecord(ByVal RecordTxt As String)
Dim Cnt As Long, Cnt2 As Integer
If RecScroll.Value < RNum Then
    For Cnt = (RecScroll.Value + 1) To RNum
        For Cnt2 = 1 To 12
            If InStr(LCase(Trim(DataBase(Cnt, Cnt2))), RecordTxt) > 0 Then
                LoadRecord Cnt
                Exit Sub
            End If
        Next Cnt2
    Next Cnt
    
    MsgBox "No More Results!"
End If
End Sub


Private Sub LoadData()
Dim FNum As Long, LNum As Long, RecData As String * 305
FNum = FreeFile

Open (App.Path + "\Database.dat") For Binary Access Read As FNum
LNum = LOF(FNum): RNum = 0

While EOF(FNum) = False
    If LNum >= 305 Then
        RNum = RNum + 1
        Get #FNum, , RecData
        ProcessData RecData
    Else
        Get #FNum, , RecData
    End If
Wend

RecScroll.Max = RNum

Close #FNum
End Sub

Private Sub ProcessData(ByVal StringData As String)
DataBase(RNum, 1) = Mid(StringData, 2, Asc(Left(StringData, 1)))
DataBase(RNum, 2) = Mid(StringData, 18, Asc(Mid(StringData, 17, 1)))
DataBase(RNum, 3) = Mid(StringData, 34, Asc(Mid(StringData, 33, 1)))
DataBase(RNum, 4) = Mid(StringData, 143, Asc(Mid(StringData, 142, 1)))
DataBase(RNum, 5) = Mid(StringData, 154, Asc(Mid(StringData, 153, 1)))
DataBase(RNum, 6) = Mid(StringData, 163, Asc(Mid(StringData, 162, 1)))
DataBase(RNum, 7) = Mid(StringData, 50, Asc(Mid(StringData, 49, 1)))
DataBase(RNum, 8) = Mid(StringData, 169, Asc(Mid(StringData, 168, 1)))
DataBase(RNum, 9) = Mid(StringData, 215, Asc(Mid(StringData, 214, 1)))
DataBase(RNum, 10) = Mid(StringData, 261, Asc(Mid(StringData, 260, 1)))
DataBase(RNum, 11) = Mid(StringData, 66, Asc(Mid(StringData, 65, 1)))
DataBase(RNum, 12) = Mid(StringData, 122, Asc(Mid(StringData, 121, 1)))
End Sub

Private Sub Form_Activate()
txtSearch.SetFocus
End Sub
