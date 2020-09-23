VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{683364A1-B37D-11D1-ADC5-006008A5848C}#1.0#0"; "dhtmled.ocx"
Begin VB.Form frmMailSender 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DM Easy Mail Sender Beta 1"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7065
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   7065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DHTMLEDLibCtl.DHTMLEdit DHTML 
      Height          =   2670
      Left            =   -15
      TabIndex        =   13
      Top             =   2325
      Width           =   7080
      ActivateApplets =   0   'False
      ActivateActiveXControls=   0   'False
      ActivateDTCs    =   -1  'True
      ShowDetails     =   0   'False
      ShowBorders     =   0   'False
      Appearance      =   1
      Scrollbars      =   -1  'True
      ScrollbarAppearance=   1
      SourceCodePreservation=   -1  'True
      AbsoluteDropMode=   0   'False
      SnapToGrid      =   0   'False
      SnapToGridX     =   50
      SnapToGridY     =   50
      BrowseMode      =   0   'False
      UseDivOnCarriageReturn=   0   'False
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   255
      Top             =   5910
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   22
      ImageHeight     =   21
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0A28
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":100E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":15F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1BDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":21C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":27A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2D8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3372
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3958
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3F3E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4524
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar2 
      Height          =   330
      Left            =   30
      TabIndex        =   12
      Top             =   1890
      Width           =   3930
      _ExtentX        =   6932
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Style           =   1
      ImageList       =   "ImageList2"
      DisabledImageList=   "ImageList2"
      HotImageList    =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   12
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Bold"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Italic"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Underline"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Align Left"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Center"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Align Right"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Fonts"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Insert Image"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Insert Hyperlink"
            ImageIndex      =   8
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   810
      Top             =   5925
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4B0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4C1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4D2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4E40
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4F52
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":5064
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":5176
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":5288
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":539A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picstatus 
      BackColor       =   &H00D6DFDE&
      Height          =   300
      Left            =   0
      ScaleHeight     =   240
      ScaleWidth      =   7020
      TabIndex        =   10
      Top             =   5040
      Width           =   7080
      Begin VB.Image imgstatus 
         Height          =   240
         Left            =   15
         Picture         =   "Form1.frx":54AC
         Top             =   15
         Width           =   240
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status : Ide"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   360
         TabIndex        =   11
         Top             =   30
         Width           =   810
      End
   End
   Begin Project1.Bevel Bevel3 
      Height          =   390
      Left            =   30
      TabIndex        =   9
      Top             =   1470
      Width           =   7020
      _ExtentX        =   12383
      _ExtentY        =   688
   End
   Begin VB.PictureBox picbar 
      BackColor       =   &H00636163&
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   45
      ScaleHeight     =   360
      ScaleWidth      =   6990
      TabIndex        =   6
      Top             =   1485
      Width           =   6990
      Begin VB.TextBox txtMail 
         Height          =   285
         Index           =   2
         Left            =   840
         TabIndex        =   8
         Top             =   30
         Width           =   5970
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Subject :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   2
         Left            =   105
         TabIndex        =   7
         Top             =   75
         Width           =   630
      End
   End
   Begin Project1.Bevel Bevel1 
      Height          =   900
      Left            =   30
      TabIndex        =   5
      Top             =   540
      Width           =   7020
      _ExtentX        =   12383
      _ExtentY        =   1588
   End
   Begin VB.TextBox txtMail 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   0
      Left            =   630
      TabIndex        =   2
      Top             =   660
      Width           =   6225
   End
   Begin VB.TextBox txtMail 
      Height          =   285
      Index           =   1
      Left            =   630
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1005
      Width           =   6225
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   435
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7065
      _ExtentX        =   12462
      _ExtentY        =   767
      ButtonWidth     =   767
      ButtonHeight    =   714
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      DisabledImageList=   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   13
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Send Mail"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Open Mail"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Save Mail"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cut"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Copy"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Paste"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Config"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "About"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exit.."
            ImageIndex      =   11
         EndProperty
      EndProperty
   End
   Begin MSWinsockLib.Winsock smtp 
      Left            =   6390
      Top             =   6225
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   1
      Left            =   885
      Picture         =   "Form1.frx":5836
      Top             =   5475
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   0
      Left            =   615
      Picture         =   "Form1.frx":5BC0
      Top             =   5490
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "From :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   1
      Left            =   180
      TabIndex        =   4
      Top             =   1020
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "To :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   195
      TabIndex        =   3
      Top             =   690
      Width           =   270
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   -195
      X2              =   1080
      Y1              =   465
      Y2              =   465
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   -195
      X2              =   1080
      Y1              =   450
      Y2              =   450
   End
End
Attribute VB_Name = "frmMailSender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Response As String
Dim MailBody As String
Dim mIndex As Integer
Dim HTML_BODY As String
Sub WaitFor(ResponseCode As String)
    ' This code in this function was not writen by me just found on the net
    ' But just like to say thank's to who ever did write it.
    
    start = Timer ' Time event so won't get stuck in loop
    While Len(Response) = 0
        Tmr = start - Timer
        DoEvents ' Let System keep checking for incoming response **IMPORTANT**
        If Tmr > 50 Then ' Time in seconds to wait
            MsgBox "SMTP service error, timed out while waiting for response", 64
            Exit Sub
        End If
    Wend
    While Left(Response, 3) <> ResponseCode
        DoEvents
        If Tmr > 50 Then
            MsgBox "SMTP service error, impromper response code. Code should have been: " + ResponseCode + " Code recieved: " + Response, 64
            Exit Sub
        End If
    Wend

Response = "" ' Sent response code to blank **IMPORTANT**
End Sub
Sub SendEmail()
Dim MailBody As String
On Error Resume Next
    smtp.Close
    smtp.LocalPort = 0
    smtp.Protocol = sckTCPProtocol
    smtp.RemoteHost = TMail.MailServer
    smtp.RemotePort = TMail.MailServerPort
    smtp.Connect
    imgstatus.Picture = Image1(1).Picture
    WaitFor ("220")
    lblStatus.Caption = "Status: Connecting to " & smtp.RemoteHost
    Replay "HELO " & smtp.LocalHostName
    WaitFor ("250")
    lblStatus.Caption = "Status: Sending mail message"
    '
    Replay "MAIL FROM: " & TMail.MailFrom
    WaitFor ("250")
    Replay "RCPT TO: " & TMail.MailTo
    WaitFor ("250")
    Replay "DATA"
    WaitFor ("354")
    Replay TMail.MailBody
    WaitFor ("250")
    lblStatus.Caption = "Status: Mail message sent"
    Replay "QUIT"
    lblStatus.Caption = "Status: Closing connection."
    WaitFor ("221")
    smtp.Close
    imgstatus.Picture = Image1(0).Picture
    lblStatus.Caption = "Status : Ide"
    
End Sub

Sub Replay(StrBuff As String)
    If smtp.State = sckConnected Then
        smtp.SendData StrBuff & vbCrLf
    End If
    
End Sub

Private Sub DHTML_onclick()
    mIndex = 3
    Toolbar1.Buttons(6).Enabled = True
    Toolbar1.Buttons(7).Enabled = True
    Toolbar1.Buttons(8).Enabled = True
    Toolbar1.Buttons(6).Image = ImageList1.ListImages(4).Index
    Toolbar1.Buttons(7).Image = ImageList1.ListImages(5).Index
    Toolbar1.Buttons(8).Image = ImageList1.ListImages(6).Index
    
End Sub

Private Sub Form_Load()
    FlatBorder txtMail(0).hwnd
    FlatBorder txtMail(1).hwnd
    FlatBorder txtMail(2).hwnd
    FlatBorder picstatus.hwnd
    Bevel1.BevelSytle VbRaised
    If FindFile(AddBackSlash(App.Path) & "config.ini") = False Then
        FirstTimeLoad = True
        frmMailSender.Hide
        MsgBox "Your new mail sender needs to be setup before you can use it.", vbInformation, "Setup Mail Sender"
        frmConfig.Show
        Exit Sub
    Else
        frmMailSender.Show
        txtMail(1).Text = ReadConfig("DM-EASYMAIL", "Mailtext")
        TMail.MailServer = ReadConfig("DM-EASYMAIL", "Servername")
        TMail.MailServerPort = Val(ReadConfig("DM-EASYMAIL", "Serverport"))
        TMail.MailFrom = ReadConfig("DM-EASYMAIL", "Mailtext")
    End If
    DHTML.SetFocus
    
End Sub

Private Sub Form_Resize()
    Line1.X2 = frmMailSender.ScaleWidth
    Line2.X2 = frmMailSender.ScaleWidth
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Response = ""
    MailBody = ""
    mIndex = 0
    Set frmMailSender = Nothing
    Set frmabout = Nothing
    Set frmConfig = Nothing
    Unload frmMailSender
    Unload frmabout
    Unload frmConfig
    
End Sub

Private Sub smtp_DataArrival(ByVal bytesTotal As Long)
    smtp.GetData Response

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim FileExt As String, lzFileName As String, MailData As String
Dim FileIn As Long, FileOut As Long
Dim ipart As Integer, lpart As Integer

    FileOut = FreeFile
    FileIn = FreeFile
    
    Select Case Button.Index
        Case 2
            On Error Resume Next
            TMail.MailTo = Trim(txtMail(0).Text)
            TMail.Subject = Trim(txtMail(2).Text)
            TMail.StrDate = Format(Now, "ddd, dd mmm yyyy hh:mm:ss  +0200")
            TMail.MailMess = DHTML.DocumentHTML
    
            If Len(TMail.MailTo) <= 0 Then
                MsgBox "The mail could not be send please include a recipient name.", vbCritical, "Error"
                Exit Sub
            End If
            If isVaildEmail(TMail.MailTo) = False Then
                MsgBox "You have not enter a invalid email address the mail will not be sent.", vbCritical, "Inviald E-Mail Address"
                txtMail(0).SetFocus
                txtMail(0).SelStart = 0
                txtMail(0).SelLength = Len(txtMail(0))
                Exit Sub
            End If
            If Len(TMail.Subject) <= 0 Then
                TMail.Subject = "No Subject...."
            End If
            
    ' Main mail message body
            TMail.MailBody = "Date: " & TMail.StrDate & vbCrLf _
            & "From: " & Mid(TMail.MailFrom, 1, InStr(TMail.MailFrom, "@") - 1) & " " & "<" & TMail.MailFrom & ">" & vbCrLf _
            & "X-Mailer: Dm Mail Sender V1.1" & vbCrLf _
            & "X-Accept-Language: en" & vbCrLf _
            & "MIME-Version: 1.0" & vbCrLf _
            & "To: " & TMail.MailTo & vbCrLf _
            & "Subject: " & TMail.Subject & vbCrLf _
            & "Content-Type: text/html;" & vbCrLf _
            & vbTab & "charset=" & Chr(34) & "iso-8859-1" & Chr(34) & vbCrLf _
            & "Content-Transfer-Encoding: 7bit" & vbCrLf _
            & vbCrLf & TMail.MailMess & vbCrLf & "."
            SendEmail
            
        Case 3
            lzFileName = OpenFile(frmMailSender.hwnd)
            If Len(lzFileName) <= 0 Then Exit Sub
            FileExt = UCase(Right(lzFileName, 3))
            If Not FileExt = "DME" Then
                MsgBox "This is not a valid DM Easy Mail document.", vbCritical, "Error"
                Exit Sub
            Else
                Open lzFileName For Binary As #FileIn
                    MailData = Space(LOF(FileIn))
                    Get #FileIn, , MailData
                Close #FileIn
                ipart = InStr(MailData, "[MailFrom]")
                lpart = InStr(ipart, MailData, vbCrLf)
                txtMail(0).Text = Mid(MailData, ipart + 10, lpart - ipart - 10)
            
                ipart = InStr(MailData, "[MAILTO]")
                lpart = InStr(ipart, MailData, vbCrLf)
                txtMail(1).Text = Mid(MailData, ipart + 8, lpart - ipart - 8)
                
                ipart = InStr(MailData, "[SUBJECT]")
                lpart = InStr(ipart, MailData, vbCrLf)
                txtMail(2).Text = Mid(MailData, ipart + 9, lpart - ipart - 9)
                    
                ipart = InStr(MailData, "[MAILDATA]")
                lpart = InStr(ipart, MailData, Chr(25))
                DHTML.DocumentHTML = Mid(MailData, ipart + 10, lpart - ipart - 10)
                
                MailData = ""
                lzFileName = ""
                FileExt = ""
                ipart = 0
                lpart = 0
            End If
            
        Case 4
            txtMail(0).Text = Trim(txtMail(0).Text)
            txtMail(1).Text = Trim(txtMail(1).Text)
            txtMail(2).Text = Trim(txtMail(2).Text)
            
            If Len(txtMail(0).Text) <= 0 Then
                MsgBox "The mail could not be saved No Mail From has been completed.", vbCritical, "Error"
                Exit Sub
            ElseIf Len(txtMail(1).Text) <= 0 Then
                MsgBox "The mail message could not be saved No Recipient has been completed.", vbCritical, "Error"
                Exit Sub
            ElseIf Len(txtMail(2).Text) <= 0 Then
                MsgBox "The mail message could not be saved No Subject has been completed.", vbCritical, "Error"
                Exit Sub
            Else
                lzFileName = SaveFile(frmMailSender.hwnd)
                FileExt = UCase(Right(lzFileName, 3))
                If Not FileExt = "DME" Then lzFileName = lzFileName & ".dme"
                If Len(FileExt) <= 0 Then Exit Sub
                
                Open lzFileName For Output As #FileOut
                    Print #FileOut, "[DM Easy Email Sender Beta 1 Do not edit below this line]"
                    Print #FileOut, "[MailFrom]" & txtMail(0).Text
                    Print #FileOut, "[MAILTO]" & txtMail(1).Text
                    Print #FileOut, "[SUBJECT]" & txtMail(2).Text
                    Print #FileOut, "[MAILDATA]" & DHTML.DocumentHTML & Chr(25)
                Close #FileOut
            End If
    
        Case 6
            If mIndex = 3 Then
                DHTML.ExecCommand DECMD_CUT
            Else
                Clipboard.SetText txtMail(mIndex).SelText
                txtMail(mIndex).SelText = ""
            End If
            
        Case 7
            If mIndex = 3 Then
                DHTML.ExecCommand DECMD_COPY
            Else
                Clipboard.SetText txtMail(mIndex).SelText
            End If
        Case 8
            If mIndex = 3 Then
                DHTML.ExecCommand DECMD_PASTE
            Else
                txtMail(mIndex).SelText = Clipboard.GetText
            End If
            
        Case 10
            frmConfig.Show vbModal
        Case 12
            frmabout.Show vbModal
        Case 13
            Unload frmMailSender
            
        End Select
        
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 2
            DHTML.ExecCommand DECMD_BOLD
        Case 3
            DHTML.ExecCommand DECMD_ITALIC
        Case 4
            DHTML.ExecCommand DECMD_UNDERLINE
        Case 6
            DHTML.ExecCommand DECMD_JUSTIFYLEFT
        Case 7
            DHTML.ExecCommand DECMD_JUSTIFYCENTER
        Case 8
            DHTML.ExecCommand DECMD_JUSTIFYRIGHT
        Case 10
            DHTML.ExecCommand DECMD_FONT
        Case 11
            DHTML.ExecCommand DECMD_IMAGE
        Case 12
            DHTML.ExecCommand DECMD_HYPERLINK
    End Select
        
End Sub

Private Sub txtMail_Click(Index As Integer)
    mIndex = Index
    Toolbar1.Buttons(6).Enabled = True
    Toolbar1.Buttons(7).Enabled = True
    Toolbar1.Buttons(8).Enabled = True

    Toolbar1.Buttons(6).Image = ImageList1.ListImages(4).Index
    Toolbar1.Buttons(7).Image = ImageList1.ListImages(5).Index
    Toolbar1.Buttons(8).Image = ImageList1.ListImages(6).Index
    
End Sub

Private Sub txtMail_GotFocus(Index As Integer)
    txtMail(Index).BackColor = 14073525
    
End Sub

Private Sub txtMail_LostFocus(Index As Integer)
    txtMail(Index).BackColor = vbWhite
    
End Sub
