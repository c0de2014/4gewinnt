VERSION 5.00
Begin VB.Form Main 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "4-Gewinnt"
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11865
   DrawMode        =   15  'Merge Pen Not
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   11865
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Draw 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3495
      Left            =   6840
      Picture         =   "Spielfeld.frx":0000
      ScaleHeight     =   3465
      ScaleWidth      =   4545
      TabIndex        =   16
      Top             =   1320
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.PictureBox Siegbild 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3495
      Left            =   6840
      Picture         =   "Spielfeld.frx":2A52
      ScaleHeight     =   3495
      ScaleWidth      =   4575
      TabIndex        =   13
      Top             =   1320
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.CommandButton Start 
      Caption         =   "Spiel beginnen"
      Height          =   495
      Left            =   6840
      TabIndex        =   12
      Top             =   240
      Width           =   4575
   End
   Begin VB.PictureBox rot 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   840
      Picture         =   "Spielfeld.frx":56AE
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   9
      Top             =   6840
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Gelb 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   840
      Picture         =   "Spielfeld.frx":5A62
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   8
      Top             =   5880
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Spielfeld 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4695
      Left            =   510
      Picture         =   "Spielfeld.frx":5F44
      ScaleHeight     =   4695
      ScaleWidth      =   6135
      TabIndex        =   7
      Top             =   840
      Width           =   6135
   End
   Begin VB.CommandButton Row7 
      Caption         =   "7"
      Height          =   495
      Left            =   5640
      Picture         =   "Spielfeld.frx":5F072
      TabIndex        =   6
      Top             =   240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Row6 
      Caption         =   "6"
      Height          =   495
      Left            =   4800
      TabIndex        =   5
      Top             =   240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Row5 
      Caption         =   "5"
      Height          =   495
      Left            =   3960
      TabIndex        =   4
      Top             =   240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Row4 
      Caption         =   "4"
      Height          =   495
      Left            =   3120
      TabIndex        =   3
      Top             =   240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Row3 
      Caption         =   "3"
      Height          =   495
      Left            =   2280
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Row2 
      Caption         =   "2"
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Row1 
      Caption         =   "1"
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Unentschieden 
      Caption         =   "Unentschieden..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2160
      TabIndex        =   17
      Top             =   6240
      Visible         =   0   'False
      Width           =   5775
   End
   Begin VB.Label Sieg2 
      Caption         =   "Spieler 2 hat gewonnen!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2040
      TabIndex        =   15
      Top             =   6240
      Visible         =   0   'False
      Width           =   8415
   End
   Begin VB.Label Sieg1 
      Caption         =   "Spieler 1 hat gewonnen!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2040
      TabIndex        =   14
      Top             =   6240
      Visible         =   0   'False
      Width           =   8415
   End
   Begin VB.Label Spieler2 
      Caption         =   "Spieler 2 ist dran..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   11
      Top             =   6960
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.Label Spieler1 
      Caption         =   "Spieler 1 ist dran..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   10
      Top             =   6000
      Visible         =   0   'False
      Width           =   3735
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim feld(7, 6) As Integer
Dim spieler As Integer

Sub leeren()

    Dim i As Integer
    Dim j As Integer
    
    For j = 0 To 6 Step 1
        For i = 0 To 7 Step 1
            feld(i, j) = 0
        Next i
    Next j

End Sub


Sub ausgeben()

    Dim i As Integer
    Dim j As Integer
    Dim Sieg As Integer
    
    
    
    For i = 6 To 1 Step -1
        For j = 1 To 7 Step 1
            If feld(j, i) = 1 Then
                Spielfeld.PaintPicture Gelb, (120 + ((j - 1) * 835)), (120 + ((i - 1) * 720))
            ElseIf feld(j, i) = 2 Then
                Spielfeld.PaintPicture rot, (120 + ((j - 1) * 835)), (120 + ((i - 1) * 720))
            End If
        Next j
    Next i
    
    Sieg = gewonnen()
    
    If Sieg = 1 Or Sieg = 2 Then
        
        Row1.Visible = False
        Row2.Visible = False
        Row3.Visible = False
        Row4.Visible = False
        Row5.Visible = False
        Row6.Visible = False
        Row7.Visible = False
        Siegbild.Visible = True
        
        Gelb.Visible = False
        Spieler1.Visible = False
        rot.Visible = False
        Spieler2.Visible = False
        
        If Sieg = 1 Then
            Sieg1.Visible = True
        Else
            Sieg2.Visible = True
        End If
    ElseIf Sieg = 3 Then
        
        Row1.Visible = False
        Row2.Visible = False
        Row3.Visible = False
        Row4.Visible = False
        Row5.Visible = False
        Row6.Visible = False
        Row7.Visible = False
        Draw.Visible = True
        
        Gelb.Visible = False
        Spieler1.Visible = False
        rot.Visible = False
        Spieler2.Visible = False
        Draw.Visible = True
        Unentschieden.Visible = True
        
    Else
        If spieler = 1 Then
            Gelb.Visible = True
            Spieler1.Visible = True
            rot.Visible = False
            Spieler2.Visible = False
        ElseIf spieler = 2 Then
            Gelb.Visible = False
            Spieler1.Visible = False
            rot.Visible = True
            Spieler2.Visible = True
        Else
            Gelb.Visible = False
            Spieler1.Visible = False
            rot.Visible = False
            Spieler2.Visible = False
        End If
    End If

End Sub

Private Sub Row1_Click()

       If feld(1, 0) < 6 And (spieler = 1 Or spieler = 2) Then
          feld(1, 6 - feld(1, 0)) = spieler
          feld(1, 0) = feld(1, 0) + 1


            If spieler = 1 Then
             spieler = 2
            ElseIf spieler = 2 Then
             spieler = 1
            End If
       End If
       ausgeben
                  
End Sub

Private Sub Row2_Click()
       
       If feld(2, 0) < 6 And (spieler = 1 Or spieler = 2) Then
          feld(2, 6 - feld(2, 0)) = spieler
          feld(2, 0) = feld(2, 0) + 1

    
           If spieler = 1 Then
            spieler = 2
           ElseIf spieler = 2 Then
            spieler = 1
           End If

       End If
       ausgeben
                  
End Sub

Private Sub Row3_Click()
       
       If feld(3, 0) < 6 And (spieler = 1 Or spieler = 2) Then
          feld(3, 6 - feld(3, 0)) = spieler
          feld(3, 0) = feld(3, 0) + 1

    
           If spieler = 1 Then
            spieler = 2
           ElseIf spieler = 2 Then
            spieler = 1
           End If

       End If
       ausgeben
                  
End Sub

Private Sub Row4_Click()
       
       If feld(4, 0) < 6 And (spieler = 1 Or spieler = 2) Then
          feld(4, 6 - feld(4, 0)) = spieler
          feld(4, 0) = feld(4, 0) + 1

    
           If spieler = 1 Then
            spieler = 2
           ElseIf spieler = 2 Then
            spieler = 1
           End If

       End If
       ausgeben
                  
End Sub

Private Sub Row5_Click()
       
       If feld(5, 0) < 6 And (spieler = 1 Or spieler = 2) Then
          feld(5, 6 - feld(5, 0)) = spieler
          feld(5, 0) = feld(5, 0) + 1


            If spieler = 1 Then
             spieler = 2
            ElseIf spieler = 2 Then
             spieler = 1
            End If
       End If

       ausgeben
                  
End Sub

Private Sub Row6_Click()
       
       If feld(6, 0) < 6 And (spieler = 1 Or spieler = 2) Then
          feld(6, 6 - feld(6, 0)) = spieler
          feld(6, 0) = feld(6, 0) + 1


            If spieler = 1 Then
             spieler = 2
            ElseIf spieler = 2 Then
             spieler = 1
            End If
       End If

       ausgeben
                  
End Sub

Private Sub Row7_Click()
       
       If feld(7, 0) < 6 And (spieler = 1 Or spieler = 2) Then
          feld(7, 6 - feld(7, 0)) = spieler
          feld(7, 0) = feld(7, 0) + 1
          
          If spieler = 1 Then
             spieler = 2
          ElseIf spieler = 2 Then
             spieler = 1
          End If
       
       End If
       
       ausgeben
                  
End Sub

Function gewonnen() As Integer
    
    Dim i As Integer
    Dim j As Integer
    Dim fertig As Boolean
    
    
    gewonnen = 0
    fertig = False
        
    'prüft ob 4 waagrecht
    For j = 1 To 6 Step 1
        For i = 1 To 4 Step 1
            If feld(i, j) = 1 And feld(i + 1, j) = 1 And feld(i + 2, j) = 1 And feld(i + 3, j) = 1 Then
                gewonnen = 1
                fertig = True
            ElseIf feld(i, j) = 2 And feld(i + 1, j) = 2 And feld(i + 2, j) = 2 And feld(i + 3, j) = 2 Then
                gewonnen = 2
                fertig = True
            End If
        Next i
    Next j
    
    'prüft ob 4 senkrecht
    If Not fertig Then
        For j = 1 To 3 Step 1
            For i = 1 To 7 Step 1
                If feld(i, j) = 1 And feld(i, j + 1) = 1 And feld(i, j + 2) = 1 And feld(i, j + 3) = 1 Then
                    gewonnen = 1
                    fertig = True
                ElseIf feld(i, j) = 2 And feld(i, j + 1) = 2 And feld(i, j + 2) = 2 And feld(i, j + 3) = 2 Then
                    gewonnen = 2
                    fertig = True
                End If
            Next i
        Next j
    End If
    
    'prüft ob 4 diagonal rechts
    If Not fertig Then
        For j = 1 To 3 Step 1
            For i = 1 To 4 Step 1
                If feld(i, j) = 1 And feld(i + 1, j + 1) = 1 And feld(i + 2, j + 2) = 1 And feld(i + 3, j + 3) = 1 Then
                    gewonnen = 1
                    fertig = True
                ElseIf feld(i, j) = 2 And feld(i + 1, j + 1) = 2 And feld(i + 2, j + 2) = 2 And feld(i + 3, j + 3) = 2 Then
                    gewonnen = 2
                    fertig = True
                End If
            Next i
        Next j
    End If
    
    'prüft ob 4 diagonal links
    If Not fertig Then
        For j = 4 To 6 Step 1
            For i = 1 To 4 Step 1
                If feld(i, j) = 1 And feld(i + 1, j - 1) = 1 And feld(i + 2, j - 2) = 1 And feld(i + 3, j - 3) = 1 Then
                    gewonnen = 1
                    fertig = True
                ElseIf feld(i, j) = 2 And feld(i + 1, j - 1) = 2 And feld(i + 2, j - 2) = 2 And feld(i + 3, j - 3) = 2 Then
                    gewonnen = 2
                    fertig = True
                End If
            Next i
        Next j
    End If
    
    If Not fertig And feld(1, 0) = 6 And feld(2, 0) = 6 And feld(3, 0) = 6 And feld(4, 0) = 6 And feld(5, 0) = 6 And feld(6, 0) = 6 And feld(7, 0) = 6 Then
        gewonnen = 3
    End If
    
End Function

Private Sub Start_Click()

    leeren

    Start.Visible = False
    Row1.Visible = True
    Row2.Visible = True
    Row3.Visible = True
    Row4.Visible = True
    Row5.Visible = True
    Row6.Visible = True
    Row7.Visible = True

    Sieg1.Visible = False
    Sieg2.Visible = False
    Siegbild.Visible = False
    
    spieler = 1
    
    ausgeben
    
End Sub
