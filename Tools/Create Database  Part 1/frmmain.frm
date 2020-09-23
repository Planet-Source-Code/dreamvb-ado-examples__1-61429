VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmmain 
   Caption         =   "Create Empty Database"
   ClientHeight    =   1530
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6405
   LinkTopic       =   "Form1"
   ScaleHeight     =   1530
   ScaleWidth      =   6405
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdexit 
      Caption         =   "E&xit"
      Height          =   330
      Left            =   3060
      TabIndex        =   4
      Top             =   900
      Width           =   1050
   End
   Begin MSComDlg.CommonDialog CDLG 
      Left            =   5760
      Top             =   1095
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
   Begin VB.CommandButton cmdCreate 
      Caption         =   "Create Db"
      Enabled         =   0   'False
      Height          =   330
      Left            =   1155
      TabIndex        =   0
      Top             =   900
      Width           =   1635
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
' This example shows how to
' Create new blank database

' See part two for next example on adding a new table and fields and inserting data

Public Function IsFileHere(lzFileName As String) As Boolean
    ' We use this to check if a file is found
    If Dir(lzFileName) = "" Then IsFileHere = False: Exit Function Else IsFileHere = True
End Function

Private Function CreateDataBase(lpFileName As String) As Boolean
Dim db_Cat As ADOX.Catalog
Dim Provider As String
Dim Data_Src As String
    
    Set db_Cat = New ADOX.Catalog ' Create the Catalog to create the database
    Provider = "Provider=Microsoft.Jet.OLEDB.4.0;" 'Provider
    Data_Src = "Data Source=" & lpFileName & ";" ' Data Source of the database
    db_Cat.Create Provider & Data_Src ' Create the new blank database
    Set db_Cat = Nothing ' Destroy the ADOX object
    CreateDataBase = True ' send back good value
    ' Clear up used vars
    Provider = ""
    Data_Src = ""
    
    Exit Function
    
CreateERR:
    ' Database was not created
    If Err Then
        MsgBox "Error: " & vbCrLf & "Number: " & Err.Number & vbCrLf _
        & "Source: " & Err.Source & vbCrLf & "Description: " & Err.Description _
        & vbCrLf & "Error Found at: Sub ListUsers", vbInformation
    End If
    
    CreateDataBase = False
End Function

Private Sub cmdCreate_Click()
Dim Created As Boolean
Dim Ans As Integer

    If IsFileHere(txtSource.Text) Then
        Ans = MsgBox("A database with that name is already present." _
        & vbCrLf & "Do you want to replace this file with the new one?", vbYesNo Or vbInformation)
        If Ans = vbNo Then Exit Sub 'Stop and do nothing
        Kill txtSource.Text  ' Delete the old filename first
        If Err Then MsgBox Err.Description, vbExclamation, "Error": Exit Sub ' show error message if have an error
    End If
    
    Created = CreateDataBase(txtSource.Text)  ' Call CreateDataBase
    
    If Not Created Then
        MsgBox "Database was not created.", vbCritical, frmmain.Caption
    Else
        MsgBox "The Database has now been created." & _
        vbCrLf & "Saved to : " & vbCrLf & txtSource.Text, vbInformation, frmmain.Caption
    End If
    
End Sub

Private Sub cmdexit_Click()
    Unload frmmain
End Sub

Private Sub cmdsource_Click()
On Error GoTo DlgErr:

    With CDLG
        txtSource.Text = ""
        .CancelError = True
        .FileName = ""
        .DialogTitle = "Create new Database"
        .Filter = "Microsoft Database(*.mdb)|*.mdb|"
        .ShowSave
        If Len(.FileName) = 0 Then Exit Sub
        txtSource.Text = .FileName
    End With
    
    Exit Sub
    
DlgErr:
    If Err = cdlCancel Then Err.Clear
    
End Sub

Private Sub txtSource_Change()
    ' Make sure that the textbox has data before enableing the cmdCreate button
    cmdCreate.Enabled = Len(txtSource.Text)
End Sub
