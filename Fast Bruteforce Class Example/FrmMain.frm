VERSION 5.00
Begin VB.Form FrmMain 
   Caption         =   "Fast BruteForce Class Example"
   ClientHeight    =   2985
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2985
   ScaleWidth      =   6000
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tRuntime 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5040
      Top             =   1800
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4800
      TabIndex        =   6
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   4800
      TabIndex        =   7
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   495
      Left            =   4800
      TabIndex        =   2
      Top             =   240
      Width           =   1095
   End
   Begin VB.Frame Frame3 
      Caption         =   "Text to crack"
      Height          =   975
      Left            =   2040
      TabIndex        =   18
      Top             =   120
      Width           =   2655
      Begin VB.TextBox txtCurrent 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   960
         TabIndex        =   21
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtText 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Generated:"
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   600
         Width           =   795
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Options"
      Height          =   1695
      Left            =   2040
      TabIndex        =   14
      Top             =   1200
      Width           =   2655
      Begin VB.TextBox txtComboLen 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         TabIndex        =   5
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox txtStartCombo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox txtCharacterSet 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Text            =   "abcdefghijklmnopqrstuvwxyz"
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Start Length:"
         Height          =   195
         Left            =   1440
         TabIndex        =   17
         Top             =   960
         Width           =   915
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Starting String:"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   960
         Width           =   1035
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Character Set:"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   1020
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Stats"
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
      Begin VB.TextBox txtSL 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   20
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox txtRuntime 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   13
         Top             =   2280
         Width           =   1575
      End
      Begin VB.TextBox txtTC 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox txtCPS 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "String Length:"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   1440
         Width           =   990
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Running Time:"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   2040
         Width           =   1035
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Total Combinations:"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   840
         Width           =   1395
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Combinations Per Sec:"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1605
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------------
'=====================================================================================
'-------------------------------------------------------------------------------------
                
                'Date:         7:56 PM 5/09/2003
                'Progammer:    Jake Paternoster (§e7eN)
                'Email:        Hate_114@hotmail.com
                'Program Name: Fast BruteForce Class Example
                
                'Description:  By request here is an example on how to
                '              use the BruteForce Class. This Program
                '              will do a plain text crack against the
                '              text specified. This code is currently
                '              fastest bruteforce code on PSC at over
                '              20,000 combinations per second running
                '              on a Celeron 900.
                '
                '              Please remember to Vote and Comment
                
'-------------------------------------------------------------------------------------
'=====================================================================================
'-------------------------------------------------------------------------------------


Public WithEvents cBF As clsBF
Attribute cBF.VB_VarHelpID = -1
Dim lRunningTime As Long
Dim bCrack As Boolean

Private Sub cBF_CombinationsPerSec(Combos As Long)
    txtCPS.Text = Combos
    txtCurrent.Text = cBF.CurrentPassword
    txtSL.Text = Len(txtCurrent.Text)
End Sub

Private Sub cBF_TotalCombinations(Combos As String)
    txtTC.Text = Combos
End Sub

Private Sub cmdExit_Click()
End
End Sub

Private Sub cmdStart_Click()
Dim sTmp As String

lRunningTime = 0
bCrack = True

DisEnableControls

With cBF
    .CharacterSet = txtCharacterSet.Text
    .FirstPassword = txtStartCombo.Text
    If txtComboLen.Text <> "" Then .StartLength = CInt(txtComboLen.Text)
    .Initialize

Do Until bCrack = False Or cBF.CurrentPassword = txtText.Text
    DoEvents
    sTmp = .BruteForce
Loop

If bCrack = True Then
    txtCurrent.Text = .CurrentPassword
    MsgBox "'" & txtText.Text & "' Cracked in " & txtRuntime.Text, vbApplicationModal + vbInformation, Me.Caption
End If

bCrack = False
DisEnableControls

End With
End Sub

Sub DisEnableControls()
    tRuntime.Enabled = Not tRuntime.Enabled
    txtText.Enabled = Not txtText.Enabled
    txtCPS.Enabled = Not txtCPS.Enabled
    txtTC.Enabled = Not txtTC.Enabled
    txtCharacterSet.Enabled = Not txtCharacterSet.Enabled
    txtStartCombo.Enabled = txtStartCombo.Enabled
    txtComboLen.Enabled = Not txtComboLen.Enabled
    cmdStart.Enabled = Not cmdStart.Enabled
    cmdStop.Enabled = Not cmdStop.Enabled
End Sub

Function TimeConv(Sec As Long) As String
Dim iSeconds As Integer
Dim iMinurts As Integer
Dim iHours As Integer
Dim iDays As Integer

iSeconds = Sec Mod 60
iMinurts = Int(Sec / 60)
iHours = Int(iMinurts / 60)
iDays = Int(iHours / 24)

TimeConv = iDays & " Days " & iHours & ":" & iMinurts & ":" & iSeconds

End Function

Private Sub cmdStop_Click()
    bCrack = False
End Sub

Private Sub Form_Load()
    Set cBF = New clsBF
End Sub

Private Sub tRuntime_Timer()
    lRunningTime = lRunningTime + 1
    txtRuntime.Text = TimeConv(lRunningTime)
End Sub
