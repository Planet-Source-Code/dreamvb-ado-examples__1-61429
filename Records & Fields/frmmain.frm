VERSION 5.00
Begin VB.Form frmmain 
   Caption         =   "Records"
   ClientHeight    =   3495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5625
   LinkTopic       =   "Form1"
   ScaleHeight     =   3495
   ScaleWidth      =   5625
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   435
      Left            =   1950
      TabIndex        =   6
      Top             =   2940
      Width           =   1635
   End
   Begin VB.TextBox txtOutput 
      Height          =   1965
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   870
      Width           =   5235
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show Records"
      Height          =   435
      Left            =   135
      TabIndex        =   0
      Top             =   2940
      Width           =   1635
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Records:"
      Height          =   195
      Left            =   135
      TabIndex        =   4
      Top             =   540
      Width           =   645
   End
   Begin VB.Label lblRecCnt 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Left            =   975
      TabIndex        =   3
      Top             =   540
      Width           =   90
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Left            =   975
      TabIndex        =   2
      Top             =   285
      Width           =   90
   End
   Begin VB.Label lblrec 
      AutoSize        =   -1  'True
      Caption         =   "Fields:"
      Height          =   195
      Left            =   135
      TabIndex        =   1
      Top             =   285
      Width           =   450
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim AdoConn As Connection
Dim AdoRecSet As Recordset

Private Sub cmdExit_Click()
    Unload frmmain ' we unload here
End Sub

Private Sub Command1_Click()
Dim RecCnt As Integer
Dim nErr As Integer
Dim StrErr As String
Dim TableName As String
Dim StrRecData As String, StrFields As String

On Error GoTo AdoError:
' Example that opens a database , show the number of records and fields
' and also lists all the records in a textbox
    
    TableName = "users" ' Table in the database to open
    Set AdoConn = New Connection
    Set AdoRecSet = New Recordset
    
    AdoConn.ConnectionString = "Data Source=..\db1.mdb" ' create connection to the database
    AdoConn.Provider = "Microsoft.Jet.OLEDB.4.0" ' set the Provider
    AdoConn.Open ' open the connection
    
    AdoRecSet.Open TableName, AdoConn, adOpenForwardOnly, adLockReadOnly ' Open Record set
    
    
    Do While Not AdoRecSet.EOF ' loop until there are no more records
        RecCnt = RecCnt + 1 ' add one to our recored counter
        ' Append all the fields and built a list of items to show in the textbox
        StrRecData = StrRecData & AdoRecSet("ID") & vbTab & AdoRecSet("FirstName") & vbTab & vbTab & AdoRecSet("LastName") _
        & vbTab & vbTab & AdoRecSet("Gender") & vbTab & AdoRecSet("Age") & vbCrLf
        AdoRecSet.MoveNext ' move to the next aviable record
    Loop
        
    StrFields = "ID" & vbTab & "FirstName:" & vbTab & "LastName:" & vbTab & "Gender:" & vbTab & "Age" & vbCrLf
    'line above adds the fields to the top of the textbox
    txtOutput.Text = StrFields & StrRecData ' show the output
    
    lblFields.Caption = AdoRecSet.Fields.Count ' show the number of Fields in the recordset
    lblRecCnt.Caption = RecCnt ' Show the total number of records
    
    AdoRecSet.Close ' close the recored set always do this first or you have an error
    AdoConn.Close ' close the connection
    ' Destroy the ado objects
    Set AdoConn = Nothing
    Set AdoRecSet = Nothing
    
    Exit Sub
    
AdoError:

    If AdoConn.Errors.Count = 0 Then Exit Sub
    ' show any errors we found
    For nErr = 0 To AdoConn.Errors.Count - 1
        StrErr = "Errors: " & AdoConn.Errors.Count & " found" & vbCrLf
        StrErr = StrErr & "Description: " & AdoConn.Errors(nErr).Description & vbCrLf
        StrErr = StrErr & "NativeError: " & AdoConn.Errors(nErr).NativeError & vbCrLf
        StrErr = StrErr & "Source: " & AdoConn.Errors(nErr).Source & vbCrLf
    Next
    
    MsgBox StrErr, vbInformation, "Ado Error" ' show the error
    
    StrErr = ""
    nErr = 0
    
    Unload frmmain
End Sub

