VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Internet Costs Controler"
   ClientHeight    =   1155
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4695
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1155
   ScaleWidth      =   4695
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAbout 
      Cancel          =   -1  'True
      Caption         =   "&About"
      Height          =   375
      Left            =   3120
      TabIndex        =   3
      Top             =   720
      Width           =   1335
   End
   Begin VB.Frame Options 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   1455
      Left            =   840
      TabIndex        =   4
      Top             =   1260
      Width           =   3015
      Begin VB.CommandButton cmdReset 
         Caption         =   "&Reset Options"
         Height          =   375
         Left            =   1560
         TabIndex        =   13
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save Options"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1200
         TabIndex        =   11
         Top             =   540
         Width           =   1695
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "Form1.frx":0442
         Left            =   2160
         List            =   "Form1.frx":04FA
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   120
         Width           =   615
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Form1.frx":05E4
         Left            =   1200
         List            =   "Form1.frx":0606
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "I&mpulse Cost:"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   540
         Width           =   945
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "s"
         Height          =   195
         Left            =   2880
         TabIndex        =   9
         Top             =   180
         Width           =   75
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "m"
         Height          =   195
         Left            =   1920
         TabIndex        =   7
         Top             =   180
         Width           =   120
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Impulse &Time:"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   180
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton cmdOptions 
      Caption         =   "&Options >>"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4080
      Top             =   120
   End
   Begin VB.Shape ProgressBar 
      DrawMode        =   6  'Mask Pen Not
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   120
      Top             =   720
      Width           =   2655
   End
   Begin VB.Shape Shape1 
      Height          =   375
      Left            =   1320
      Top             =   720
      Width           =   3255
   End
   Begin VB.Label Status 
      Alignment       =   2  'Center
      Caption         =   "You are currently Offline."
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   360
      Width           =   4455
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      BorderWidth     =   2
      X1              =   120
      X2              =   4560
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label Timed 
      Alignment       =   2  'Center
      Caption         =   "You are currently Offline."
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'General variable declaration
'ITime = impulse time
'ICost = impulse cost
'NetStart = time when you connected
'NetCount = total connection impulses
'NetCost = total connection cost
Private ITime As Integer, ICost As Single
Private NetStart As Long, NetCount As Integer, NetCost As Long
Private OnLine As Boolean

Private Sub cmdAbout_Click()
'Generate a standard About Message Box
    MsgBox "Programmed by Pedro Lamas" & vbCrLf & "Copyright ©1997-1999 Underground Software" & vbCrLf & vbCrLf & "Home-Page (Dedicated to VB): www.terravista.pt/portosanto/3723/" & vbCrLf & "E-Mail: sniper@hotpop.com", vbApplicationModal + vbInformation, "Credits!"
End Sub

Private Sub cmdClose_Click()
'End the program
    End
End Sub

Private Sub cmdOptions_Click()
    If cmdOptions.Caption = "&Options >>" Then
'Show the Options Panel
        ResetOptions
        cmdOptions.Caption = "&Options <<"
        Me.Height = 3045
        Options.Enabled = True
'Disable the timer
        Timer1.Enabled = False
    Else
'Hide the Options Panel
        cmdOptions.Caption = "&Options >>"
        Me.Height = 1545
        Options.Enabled = False
'Enable the timer
        Timer1.Enabled = True
    End If
End Sub

Private Sub cmdReset_Click()
'Reset the options, by reloading the settings
    ResetOptions
End Sub

Private Sub cmdSave_Click()
'Update the options variables
    ITime = Combo1 * 60 + Combo2
    ICost = CSng(Text1)
'Save the options on the Windows Registry
    SaveSetting "Internet Costs Control", "ops", "ITime", ITime
    SaveSetting "Internet Costs Control", "ops", "ICost", ICost
End Sub

Private Sub Form_Load()
'Set the options, by loading the setings
    ResetOptions
'Align the Impulse Progress Bar
    Shape1.Move Timed.Left, Timed.Top, Timed.Width, Timed.Height + Status.Height
    ProgressBar.Move Timed.Left + 30, Timed.Top + 30, 0, Timed.Height + Status.Height - 60
    ProgressBar.Tag = Timed.Width - 60
End Sub

Private Sub Timer1_Timer()
    On Error Resume Next
    Dim Ret As Long, PassedTime As Long
    
'Check if you are connected
    If IsRASConnected Then
'Get the time you have been connected
        PassedTime = Timer - NetStart
        If OnLine Then
'If you were already connected,
            Timed = "OnLine for " & TimeSerial(0, 0, PassedTime)
            If Int(PassedTime / ITime + 1) > NetCount Then
'Update the Costs Panel
                NetCount = Int(PassedTime / ITime + 1)
                NetCost = NetCount * ICost
                Status = "You've already spent " & Format(NetCost, "Currency")
                ProgressBar.Width = Int(ProgressBar.Tag)
            Else
'Update the Impulse Progress Bar
                ProgressBar.Width = Int(ProgressBar.Tag) / ITime * (PassedTime Mod ITime)
            End If
        Else
'If you weren't connected, set the session variables
            NetStart = Timer
            NetCount = 1
            NetCost = ICost
            Timed = "OnLine for " & TimeSerial(0, 0, 0)
            Status = "You've already spent " & Format(NetCost, "Currency")
'Change the Online flag to True
            OnLine = True
'Disable the Options CommandButton
            cmdOptions.Enabled = False
        End If
    Else
'If you were connected,
        If OnLine Then
'Update the Costs Panel
            Timed = "OffLine. You've been Online for " & TimeSerial(0, 0, Timer - NetStart)
            Status = "You've spent " & Format(NetCost, "Currency")
            ProgressBar.Width = 0
'Change the Online flag to False
            OnLine = False
'Enable the Options CommandButton
            cmdOptions.Enabled = True
        End If
    End If
'Do all events
    DoEvents
End Sub

Private Sub ResetOptions()
'Get the program settings from Windows Registry
    ITime = Int(GetSetting("Internet Costs Control", "Ops", "ITime", 180))
    ICost = CSng(GetSetting("Internet Costs Control", "Ops", "ICost", 10))
'Update the Options Controls with the setting values
    Combo1 = Int(ITime / 60)
    Combo2 = ITime - 60 * Int(ITime / 60)
    Text1 = ICost
End Sub
