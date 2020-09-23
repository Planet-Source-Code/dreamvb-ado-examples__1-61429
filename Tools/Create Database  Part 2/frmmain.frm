VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmmain 
   Caption         =   "Create New Database, add new table"
   ClientHeight    =   2130
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5790
   LinkTopic       =   "Form1"
   ScaleHeight     =   2130
   ScaleWidth      =   5790
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtTable 
      Height          =   300
      Left            =   1500
      TabIndex        =   6
      Top             =   855
      Width           =   1830
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "E&xit"
      Height          =   405
      Left            =   2265
      TabIndex        =   4
      Top             =   1485
      Width           =   1050
   End
   Begin MSComDlg.CommonDialog CDLG 
      Left            =   4650
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdsource 
      Caption         =   "...."
      Height          =   360
      Left            =   5040
      TabIndex        =   3
      Top             =   330
      Width           =   555
   End
   Begin VB.TextBox txtSource 
      Height          =   315
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   330
      Width           =   3855
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "Create db && Add Table"
      Enabled         =   0   'False
      Height          =   405
      Left            =   75
      TabIndex        =   0
      Top             =   1485
      Width           =   2070
   End
   Begin VB.Label lbltbname 
      AutoSize        =   -1  'True
      Caption         =   "New TableName:"
      Height          =   195
      Left            =   135
      TabIndex        =   5
      Top             =   930
      Width           =   1245
   End
   Begin VB.Label lblSource 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Db Source:"
      Height          =   195
      Left            =   180
      TabIndex        =   1
      Top             =   390
      Width           =   810
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' This example shows how to:

' 1. Create a new database
' 2. Add a new table
' 3. Add field names
' 4. Inserting recoreds

Dim aDox_Cat As ADOX.Catalog
Dim aDox_tb As ADOX.Table
Dim AdoConn As ADODB.Connection
Dim Table_Name As String ' this is the table in the db to load

Public Sub CleanObjects()
    If AdoConn.State = adStateOpen Then AdoConn.Close ' Close connection if it is open
    ' Destroy the objects
    Set AdoConn = Nothing
    Set aDox_Cat = Nothing
    Set aDox_tb = Nothing
End Sub

Public Function IsFileHere(lzFileName As String) As Boolean
    ' We use this to check if a file is found
    If Dir(lzFileName) = "" Then IsFileHere = False: Exit Function Else IsFileHere = True
End Function

Private Function CreateNewTable(mTableName As String) As Boolean
On Error GoTo AddArr:

    'Set the table object
    Set aDox_tb = New ADOX.Table
    
    aDox_tb.Name = mTableName ' Set the tables name
    aDox_tb.Columns.Append "Software Title", adWChar, 30 ' Add Field , DataType and the size of the data
    aDox_tb.Columns.Append "Publisher", adWChar, 30 ' Add Field , DataType and the size of the data
    aDox_Cat.Tables.Append aDox_tb ' Add the new fields and the table to the database
    
    Set AdoConn = aDox_Cat.ActiveConnection ' Set the connection for AdoConn
    
    ' Now we will use the INSERT command to insert some data to our database
    AdoConn.Execute "INSERT INTO " & mTableName & " VALUES ('Visual Basic','Microsoft')"
    AdoConn.Execute "INSERT INTO " & mTableName & " VALUES ('Delphi','Borland')"
    AdoConn.Execute "INSERT INTO " & mTableName & " VALUES ('C++ Builder Basic','Borland')"
    AdoConn.Execute "INSERT INTO " & mTableName & " VALUES ('Photo Shop','Adobe')"
    AdoConn.Execute "INSERT INTO " & mTableName & " VALUES ('VC++','Microsoft')"
    AdoConn.Execute "INSERT INTO " & mTableName & " VALUES ('Winamp','NullSoft')"
    
    CleanObjects ' Clean Up
    CreateNewTable = True ' Return good code
    Exit Function
AddArr:
    If Err Then MsgBox Err.Description, vbInformation, frmmain.Caption
    CreateNewTable = False ' Return bad code
    
End Function

Private Function CreateDataBase(lpFileName As String) As Boolean
Dim Provider As String
Dim Data_Src As String
    
    Set aDox_Cat = New ADOX.Catalog ' Create the Catalog to create the database
    Provider = "Provider=Microsoft.Jet.OLEDB.4.0;" 'Provider
    Data_Src = "Data Source=" & lpFileName & ";" ' Data Source of the database
    aDox_Cat.Create Provider & Data_Src ' Create the new blank database
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
On Error Resume Next

    If IsFileHere(txtSource.Text) Then
        Ans = MsgBox("A database with that name is already present." _
        & vbCrLf & "Do you want to replace this file with the new one?", vbYesNo Or vbInformation)
        If Ans = vbNo Then Exit Sub 'Stop and do nothing
        Kill txtSource.Text  ' Delete the old filename first
        If Err Then MsgBox Err.Description, vbExclamation, "Error": Exit Sub ' show error message if have any
    End If
    
    Created = CreateDataBase(txtSource.Text) ' Call CreateDataBase
    
    If Not Created Then
        ' Database was not created
        MsgBox "Database was not created.", vbCritical, frmmain.Caption
    ElseIf Len(Trim(txtTable.Text)) = 0 Then
        MsgBox "You need to enter a table name", vbInformation, frmmain.Caption
        Exit Sub
    Else
        Created = CreateNewTable(txtTable.Text) ' Add new table, Field names and some data
        If Not Created Then
            MsgBox "There was an error while adding the table to the database", vbInformation
            Kill txtSource.Text ' Kill the filename
            Exit Sub
        Else
            MsgBox "The database was created and the table and fields have been added." _
            & vbCrLf & vbCrLf & "Database saved to " & vbCrLf & txtSource.Text, vbInformation, frmmain.Caption
            Table_Name = ""
        End If
    End If
    
End Sub

Private Sub cmdexit_Click()
    Unload frmmain
End Sub

Private Sub cmdsource_Click()
On Error GoTo DlgErr:
Dim Ans As Integer
    With CDLG
        txtSource.Text = ""
        .CancelError = True
        .FileName = ""
        .DialogTitle = "Create new Database"
        .Filter = "Microsoft Database(*.mdb)|*.mdb|"
        .ShowSave
        If Len(.FileName) = 0 Then Exit Sub
        txtSource.Text = .FileName ' update source textbox with the database filename
    End With
    
    Exit Sub
    
DlgErr:
    If Err = cdlCancel Then Err.Clear
    
End Sub

Private Sub txtSource_Change()
    ' Make sure that the textbox has data before enableing the cmdCreate button
    cmdCreate.Enabled = Len(txtSource.Text)
End Sub
