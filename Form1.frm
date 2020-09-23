VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmURLCheck 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "URL Checker"
   ClientHeight    =   4200
   ClientLeft      =   4065
   ClientTop       =   1545
   ClientWidth     =   5655
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   5655
   Begin VB.TextBox txtResultCode 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   75
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   3660
      Width           =   825
   End
   Begin VB.Frame fraTitle 
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      Caption         =   " URL Checker "
      Height          =   990
      Left            =   -15
      TabIndex        =   10
      Top             =   0
      Width           =   5865
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Checks for broken links"
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   555
         TabIndex        =   12
         Top             =   540
         Width           =   4800
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "URL Checker"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   180
         TabIndex        =   11
         Top             =   180
         Width           =   3630
      End
   End
   Begin MSComCtl2.UpDown udnRetries 
      Height          =   375
      Left            =   5280
      TabIndex        =   9
      Top             =   2100
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   661
      _Version        =   393216
      Value           =   1
      AutoBuddy       =   -1  'True
      BuddyControl    =   "txtRetries"
      BuddyDispid     =   196613
      OrigLeft        =   5340
      OrigTop         =   1650
      OrigRight       =   5580
      OrigBottom      =   2055
      Min             =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   0
      Enabled         =   -1  'True
   End
   Begin VB.TextBox txtRetries 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Text            =   "1"
      Top             =   2100
      Width           =   600
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   375
      Left            =   3705
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2760
      Width           =   885
   End
   Begin VB.TextBox txtResult 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   945
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3660
      Width           =   4590
   End
   Begin InetCtlsObjects.Inet IE1 
      Left            =   105
      Top             =   2490
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "Check"
      Default         =   -1  'True
      Height          =   375
      Left            =   4665
      TabIndex        =   2
      Top             =   2760
      Width           =   885
   End
   Begin VB.TextBox txtURL 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   75
      TabIndex        =   0
      Top             =   1590
      Width           =   5460
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   5655
      Y1              =   3255
      Y2              =   3255
   End
   Begin VB.Label lblRetries 
      BackStyle       =   0  'Transparent
      Caption         =   "How many attempts should be made:"
      Height          =   315
      Left            =   75
      TabIndex        =   8
      Top             =   2190
      Width           =   4140
   End
   Begin VB.Label lblResults 
      BackStyle       =   0  'Transparent
      Caption         =   "Results of URL check:"
      Height          =   330
      Left            =   90
      TabIndex        =   7
      Top             =   3390
      Width           =   2895
   End
   Begin VB.Label lblURL 
      BackStyle       =   0  'Transparent
      Caption         =   "Please type in the URL that you would like to check:"
      Height          =   345
      Left            =   90
      TabIndex        =   6
      Top             =   1305
      Width           =   5550
   End
End
Attribute VB_Name = "frmURLCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCheck_Click()
'   Check the selected URL
On Error GoTo Err_cmdCheck_Click

'   Declarations
    Dim CHECK_RESULT As String

'   Check the URL
    txtResultCode.Text = ""
    txtResult.Text = "Checking"
    
    If txtURL.Text = "" Then
'       No URL entered
        txtResultCode.Text = "000"
        txtResult.Text = "No URL entered"
        
    Else
    
'       Check the URL and store the result
        CHECK_RESULT = CheckURL(txtURL.Text, Val(txtRetries.Text))
        
'       Display the result code
        txtResultCode.Text = Left$(CHECK_RESULT, 3)
        
'       Display the result description
        If Len(CHECK_RESULT) > 4 Then
            txtResult.Text = Right$(CHECK_RESULT, Len(CHECK_RESULT) - 4)
        Else
            txtResult.Text = "Error"
        End If
    
    End If

Exit_cmdCheck_Click:
    Exit Sub
    
Err_cmdCheck_Click:
'   Error Handler
    MsgBox Err.Description
    Resume Exit_cmdCheck_Click

End Sub

Private Sub cmdExit_Click()
'   Terminate the application
    End
End Sub

Private Function CheckURL(CHECK_URL As String, Optional CHECK_RETRIES As Long = 1) As String
'   This function will check a URL, if it fails it will retry for a set amount of times
On Error GoTo Err_CheckURL

'   Declarations
    Dim sResultLine As String
    Dim lRetryCounter As Long: lRetryCounter = 0
    Dim bSuccessful As Boolean: bSuccessful = False

    cmdCheck.Enabled = False

'   Loop for the selected number of retries or until successful
    Do While lRetryCounter < CHECK_RETRIES And Not (bSuccessful)

'       Add an indication to the results box to show retries
        lRetryCounter = lRetryCounter + 1
        txtResultCode.Text = lRetryCounter
        txtResult.Text = txtResult.Text + "."

'       Pass the selected URL to the INET control
        IE1.OpenURL (CHECK_URL)
    
'       Check the results
        sResultLine = Left$(IE1.GetHeader, InStr(10, IE1.GetHeader, vbCr) - 1)
        CheckURL = Right$(sResultLine, Len(sResultLine) - 9)
    
'       See if the result is a success
        If Left$(CheckURL, 3) = "200" Then
            bSuccessful = True
        End If
    
    Loop

Exit_CheckURL:
    cmdCheck.Enabled = True
    Exit Function
    
Err_CheckURL:
'   Error Handler
    MsgBox Err.Description
    Resume Exit_CheckURL
    Resume
End Function

Private Sub Form_Load()
'   Set up the form
On Error GoTo Err_Form_Load

'   Set the title banner background colour
    fraTitle.BackColor = RGB(51, 102, 204)
    
'   Force the form to be displayed in the centre of the screen
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    
Exit_Form_Load:
    Exit Sub
    
Err_Form_Load:
'   Error Handler
    MsgBox Err.Description
    Resume Exit_Form_Load
    
End Sub

Private Sub txtRetries_GotFocus()
'   Select the text
    txtRetries.SelStart = 0
    txtRetries.SelLength = Len(txtRetries.Text)
End Sub

Private Sub txtRetries_KeyUp(KeyCode As Integer, Shift As Integer)
'   Validate the number of retries
On Error GoTo Err_txtRetries_KeyUp

'   Check that the number of retries doesn't exceed 10
    If Val(txtRetries.Text) > 10 Then
        txtRetries.Text = "10"
    End If
    
'   Check that the number of retries is at least 1
    If Val(txtRetries.Text) < 1 Then
        txtRetries.Text = "1"
    End If

Exit_txtRetries_KeyUp:
    Exit Sub
    
Err_txtRetries_KeyUp:
'   Error Handler
    MsgBox Err.Description
    Resume Exit_txtRetries_KeyUp
    
End Sub

Private Sub txtURL_GotFocus()
    txtURL.SelStart = 0
    txtURL.SelLength = Len(txtURL.Text)
End Sub
