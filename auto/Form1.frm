VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8970
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8115
   LinkTopic       =   "Form1"
   ScaleHeight     =   8970
   ScaleWidth      =   8115
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Dictonary"
      Height          =   2820
      Left            =   4350
      TabIndex        =   3
      Top             =   5700
      Width           =   2880
      Begin VB.ListBox List1 
         Height          =   2400
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   300
         Width           =   2610
      End
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   2670
      Left            =   345
      TabIndex        =   2
      Top             =   1470
      Width           =   7275
      _ExtentX        =   12832
      _ExtentY        =   4710
      _Version        =   393217
      TextRTF         =   $"Form1.frx":0000
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   4770
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   870
      Width           =   2820
   End
   Begin VB.TextBox Text1 
      Height          =   555
      Left            =   375
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Form1.frx":008B
      Top             =   840
      Width           =   3330
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   $"Form1.frx":0091
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3180
      Left            =   330
      TabIndex        =   5
      Top             =   4830
      Width           =   3870
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tem
Dim rtem
Dim prevs As String
Dim tps
Dim css As Integer
Dim csl As Integer
Dim C As New Collection
Dim x As Long
Dim y As String

Private Sub Combo1_KeyPress(KeyAscii As Integer)
atcm Combo1, KeyAscii

End Sub

Private Sub RichTextBox1_KeyPress(KeyAscii As Integer)
atcm RichTextBox1, KeyAscii

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
atcm Text1, KeyAscii

End Sub
'code from ps
Function RemoveDups(LB As ListBox)


On Error Resume Next

If LB.ListCount > 1 Then
    For x = 0 To LB.ListCount - 1
        y = LB.List(x)
        C.Add y, y
    Next x
    
    LB.Clear
    
    For x = 1 To C.Count
        LB.AddItem C.Item(x)
    Next x
    
    Set C = Nothing
End If
End Function
Function atcm(ctl As Control, KeyAscii As Integer)
    If KeyAscii = 8 Then
        GoTo en
    ElseIf KeyAscii = 13 Then
        GoTo nx
    End If
    If ctl.SelLength = 0 Then
        ctl.Text = ctl.Text & Chr(KeyAscii)
    ElseIf ctl.SelLength > 0 And KeyAscii = 32 Then
        ctl.Text = ctl.Text & " "
        ctl.SelStart = Len(ctl.Text)
        KeyAscii = 0
        Exit Function
    ElseIf ctl.SelLength > 0 Then
        ctl.Text = Left(ctl.Text, Len(ctl.Text) - ctl.SelLength) & Chr(KeyAscii)
    End If
    ctl.SelStart = Len(ctl.Text)
nx:
    If KeyAscii = 32 Or KeyAscii = 13 Then
        tem = Trim(ctl.Text)
        tem = Split(tem, " ")
        If Len(tem(UBound(tem))) > 2 And UBound(tem) > 0 Then
            If InStr(tem(UBound(tem)), vbCrLf) = 0 Then
                If Len(tem(UBound(tem))) > 2 Then List1.AddItem tem(UBound(tem))
            Else
                Dim ttem
                ttem = Split(tem(UBound(tem)), vbCrLf)
               If Len(ttem(UBound(ttem))) > 2 Then List1.AddItem ttem(UBound(ttem))
            End If
        Else
            Set tem = Nothing
            tem = Split(ctl.Text, vbCrLf)
            If InStr(Trim("" & tem(UBound(tem)) & ""), " ") = 0 Then
                If Len(Trim("" & tem(UBound(tem)) & "")) > 2 Then List1.AddItem Trim("" & tem(UBound(tem)) & "")
            Else
                rtem = Split(Trim("" & tem(UBound(tem)) & ""), " ")
                If Len(rtem(UBound(rtem))) > 2 Then List1.AddItem rtem(UBound(rtem))
            End If
        End If
        RemoveDups List1
    Else
        tps = Split(ctl.Text, " ")
        If UBound(tps) = 0 Or InStr(tps(UBound(tps)), vbCrLf) > 0 Then
            Set tps = Nothing
            tps = Split(ctl.Text, vbCrLf)
        End If
        css = Len(ctl.Text)
        For i = 0 To List1.ListCount - 1
            If Len(tps(UBound(tps))) > 0 Then
                If StrComp(tps(UBound(tps)), Left(List1.List(i), Len(Trim("" & tps(UBound(tps)) & ""))), vbTextCompare) = 0 Then
                    ctl.Text = ctl.Text & Right(List1.List(i), Len(List1.List(i)) - Len(Trim("" & tps(UBound(tps)) & "")))
                    ctl.SelStart = css
                    ctl.SelLength = Len(ctl.Text) - css
                    Exit For
                End If
            Else
                Exit For
            End If
        Next i
    End If
    If KeyAscii <> 13 Then KeyAscii = 0
en:

End Function
