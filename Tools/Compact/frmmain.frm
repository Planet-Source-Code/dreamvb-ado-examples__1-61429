VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmmain 
   Caption         =   "Compact Database"
   ClientHeight    =   2070
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6405
   LinkTopic       =   "Form1"
   ScaleHeight     =   2070
   ScaleWidth      =   6405
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkEncode 
      Caption         =   "Encrypt Database"
      Height          =   270
      Left            =   2310
      TabIndex        =   7
      Top             =   1440
      Value           =   1  'Checked
      Width           =   2565
   End
   Begin VB.TextBox txtout 
      Height          =   315
      Left            =   1155
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   960
      Width           =   3855
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "...."
      Height          =   360
      Left            =   5115
      TabIndex        =   4
      Top             =   960
      Width           =   555
   End
   Begin MSComDlg.CommonDialog CDLG 
      Left            =   5880
      Top             =   285
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdsource 
      Caption         =   "...."
      Height          =   360
      Left            =   5115
      TabIndex        =   3
      Top             =   390
      Width           =   555
   End
   Begin VB.TextBox txtSource 
      Height          =   315
      Left            =   1155
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   390
      Width           =   3855
   End
   Begin VB.CommandButton cmdCompact 
      Caption         =   "Compact"
      Enabled         =   0   'False
      Height          =   330
      Left            =   450
      TabIndex        =   0
      Top             =   1410
      Width           =   1635
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Output"
      Height          =   195
      Left            =   450
      TabIndex        =   6
      Top             =   1020
      Width           =   480
   End
   Begin VB.Label lblSource 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Source:"
      Height          =   195
      Left            =   450
      TabIndex        =   1
      Top             =   450
      Width           =   555
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MSJet As JRO.JetEngine
' This example shows how to Compact a database

Private Sub cmdCompact_Click()
Dim SrcConn As String
Dim DestConn As String
On Error GoTo CompactErr:

    Set MSJet = New JRO.JetEngine ' create the Jet Object
    SrcConn = "Data Source=" & txtSource.Text ' source database
    
    DestConn = "Data Source=" & txtout.Text & ";" & _
    "Microsoft.Jet.OLEDB.4.0:Encrypt Database=" & CBool(chkEncode.Value) ' dest database
    MSJet.CompactDatabase SrcConn, DestConn ' compact the database
    MsgBox "Database has now been Compact", vbInformation, frmmain.Caption
    
    Exit Sub
    
CompactErr:
    ' Show a messagebox if we found any errors
    If Err Then
        MsgBox "Error: " & vbCrLf & "Number: " & Err.Number & vbCrLf _
        & "Source: " & Err.Source & vbCrLf & "Description: " & Err.Description _
        , vbInformation, frmmain.Caption
    End If

End Sub

Private Sub cmdOpen_Click()
On Error Resume Next

    With CDLG
        txtout.Text = ""
        .CancelError = True
        .DialogTitle = "Save"
        .Filter = "Microsoft Database(*.mdb)|*.mdb|"
        .ShowSave
        If Len(.FileName) = 0 Then Exit Sub
        
        If LCase(.FileName) = LCase(txtSource.Text) Then
            MsgBox "You cannot save to the same database." _
            & vbCrLf & "Please give a different filename.", vbExclamation, frmmain.Caption
            Exit Sub
        Else
            txtout.Text = .FileName
        End If
    End With
    
    Exit Sub
    
    If Err Then Err.Clear
End Sub

Private Sub cmdsource_Click()
On Error Resume Next

    With CDLG
        txtSource.Text = ""
        .CancelError = True
        .DialogTitle = "Open"
        .Filter = "Microsoft Database(*.mdb)|*.mdb|"
        .ShowOpen
        If Len(.FileName) = 0 Then Exit Sub
        If Not LCase(Right(.FileName, 3)) = "mdb" Then
            MsgBox "This is not a vaild Microsoft Access Database", vbInformation, frmmain.Caption
            Exit Sub
        Else
            txtSource.Text = .FileName
        End If
            
    End With
    
    Exit Sub
    
    If Err Then Err.Clear
    
End Sub

Private Sub txtout_Change()
    txtSource_Change
End Sub

Private Sub txtSource_Change()
    ' Make sure that both textboxes have data before enableing the compact button
    cmdCompact.Enabled = Len(txtSource.Text) > 0 And Len(txtout.Text) > 0
End Sub
