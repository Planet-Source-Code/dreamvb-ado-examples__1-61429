VERSION 5.00
Begin VB.Form frmmain 
   Caption         =   "Navigating"
   ClientHeight    =   2340
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4755
   LinkTopic       =   "Form1"
   ScaleHeight     =   2340
   ScaleWidth      =   4755
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   360
      Left            =   3585
      TabIndex        =   14
      Top             =   1755
      Width           =   870
   End
   Begin VB.CommandButton cmdLast 
      Caption         =   "Last"
      Height          =   360
      Left            =   2685
      TabIndex        =   13
      Top             =   1755
      Width           =   735
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   360
      Left            =   960
      TabIndex        =   12
      Top             =   1755
      Width           =   735
   End
   Begin VB.CommandButton cmdFirst 
      Caption         =   "First"
      Height          =   360
      Left            =   135
      TabIndex        =   11
      Top             =   1755
      Width           =   735
   End
   Begin VB.TextBox txtID 
      Height          =   285
      Left            =   450
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   195
      Width           =   525
   End
   Begin VB.TextBox txtAge 
      Height          =   285
      Left            =   3630
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   1170
      Width           =   525
   End
   Begin VB.TextBox txtGender 
      Height          =   285
      Left            =   3630
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   690
      Width           =   525
   End
   Begin VB.TextBox txtLastName 
      Height          =   285
      Left            =   1035
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1170
      Width           =   1725
   End
   Begin VB.TextBox txtFirstname 
      Height          =   285
      Left            =   1035
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   690
      Width           =   1725
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next"
      Height          =   360
      Left            =   1785
      TabIndex        =   0
      Top             =   1755
      Width           =   735
   End
   Begin VB.Label lblid 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ID"
      Height          =   195
      Left            =   195
      TabIndex        =   9
      Top             =   240
      Width           =   165
   End
   Begin VB.Label lblAge 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Age:"
      Height          =   195
      Left            =   2970
      TabIndex        =   6
      Top             =   1170
      Width           =   330
   End
   Begin VB.Label lblGender 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Gender:"
      Height          =   195
      Left            =   2970
      TabIndex        =   5
      Top             =   690
      Width           =   570
   End
   Begin VB.Label lblLastName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LastName:"
      Height          =   195
      Left            =   210
      TabIndex        =   2
      Top             =   1170
      Width           =   765
   End
   Begin VB.Label lblFirstName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FirstName:"
      Height          =   195
      Left            =   210
      TabIndex        =   1
      Top             =   690
      Width           =   750
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Ok so far you seen how to load a database,
' show the fields and records and also list all the records
' next we are going to play around with Navigating the databse
' and showing the data

Dim AdoConn As Connection
Dim AdoRecSet As Recordset
Dim Table_Name As String ' this is the table in the db to load

Private Sub ShowRecord()
    If AdoRecSet.EOF Then Exit Sub ' check if we are at the last record
    If AdoRecSet.BOF Then Exit Sub ' check if we are at the first record
    
    ' update the textboxes with the record data
    txtID.Text = AdoRecSet("ID")
    txtFirstname.Text = AdoRecSet("FirstName")
    txtLastName.Text = AdoRecSet("LastName")
    txtGender.Text = AdoRecSet("Gender")
    txtAge.Text = AdoRecSet("Age")
End Sub

Sub ShowAdoError()
Dim nErr As Integer
Dim StrErr As String

    ' inform of any errors
    For nErr = 0 To AdoConn.Errors.Count - 1
        StrErr = "Errors: " & AdoConn.Errors.Count & " found" & vbCrLf
        StrErr = StrErr & "Description: " & AdoConn.Errors(nErr).Description & vbCrLf
        StrErr = StrErr & "NativeError: " & AdoConn.Errors(nErr).NativeError & vbCrLf
        StrErr = StrErr & "Source: " & AdoConn.Errors(nErr).Source & vbCrLf
    Next
    
    MsgBox StrErr, vbInformation, "Ado Error" ' show the error
    
    StrErr = ""
    nErr = 0
End Sub

Sub CloseAndCleanUp()
    ' we use this sub to clean eveything we have used
    AdoRecSet.Close ' close the recored set always do this first or you have an error
    AdoConn.Close ' close the connection
    ' Destroy the ado objects
    Table_Name = ""
    Set AdoConn = Nothing
    Set AdoRecSet = Nothing
End Sub

Private Function InitializeDb() As Boolean
On Error GoTo AdoError:

    ' load the needed ado objects
    Set AdoConn = New Connection
    Set AdoRecSet = New Recordset
    
    AdoConn.ConnectionString = "Data Source=..\db1.mdb" ' create connection to the database
    AdoConn.Provider = "Microsoft.Jet.OLEDB.4.0" ' set the Provider
    AdoConn.Open ' open the connection
    InitializeDb = True
    Exit Function
    
AdoError:
    InitializeDb = False
    
End Function

Private Function LoadRecordSet() As Boolean
On Error GoTo LoadRecErr:
    AdoRecSet.Open Table_Name, AdoConn, adOpenDynamic, adLockReadOnly ' Open Record set
    LoadRecordSet = True ' send back good value
    Exit Function
    
LoadRecErr:
    ' report errors found
    MsgBox "Error:" & vbCrLf _
    & "Description: " & Err.Description _
    & vbCrLf & "Source: " & Err.Source, vbInformation, "Function LoadRecordSet()"
    LoadRecordSet = False
    
End Function

Private Sub cmdBack_Click()
    If AdoRecSet.BOF Then AdoRecSet.MoveFirst
    ' move back a record
    AdoRecSet.MovePrevious
    ShowRecord
End Sub

Private Sub cmdExit_Click()
    CloseAndCleanUp 'Call cleanup
    Unload frmmain 'unload the form
End Sub

Private Sub cmdFirst_Click()
    AdoRecSet.MoveFirst ' Move to the first record
    ShowRecord ' call ShowRecord
End Sub

Private Sub cmdLast_Click()
    AdoRecSet.MoveLast ' move to the last record
    ShowRecord ' call ShowRecord
End Sub

Private Sub cmdNext_Click()
    If AdoRecSet.EOF Then AdoRecSet.MoveLast: Exit Sub
    AdoRecSet.MoveNext ' move to the next aviable record
    ShowRecord ' call ShowRecord
End Sub

Private Sub Form_Load()
Dim InitGood As Boolean
    
    Table_Name = "users" ' table to load
    ' First thing to do is to Initialize the ado object and open the database
    InitGood = InitializeDb()
    
    If Not InitGood Then ' check for any errors
        ShowAdoError ' show error message
        Unload frmmain 'unload the form
    ElseIf Not LoadRecordSet() Then 'next we load the recordset
        Unload frmmain
    Else
        ' load in the first record and update our program
        ShowRecord
    End If
    
End Sub
