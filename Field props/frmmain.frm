VERSION 5.00
Begin VB.Form frmmain 
   Caption         =   "Fields"
   ClientHeight    =   3105
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6435
   LinkTopic       =   "Form1"
   ScaleHeight     =   3105
   ScaleWidth      =   6435
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstType 
      Height          =   1425
      Left            =   4860
      TabIndex        =   10
      Top             =   870
      Width           =   1410
   End
   Begin VB.ListBox lstDefinedSize 
      Height          =   1425
      Left            =   3300
      TabIndex        =   9
      Top             =   870
      Width           =   1410
   End
   Begin VB.ListBox LstActualSize 
      Height          =   1425
      Left            =   1755
      TabIndex        =   7
      Top             =   870
      Width           =   1410
   End
   Begin VB.ListBox lstFieldname 
      Height          =   1425
      Left            =   210
      TabIndex        =   5
      Top             =   870
      Width           =   1410
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   435
      Left            =   2040
      TabIndex        =   3
      Top             =   2460
      Width           =   1635
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show Fields/ Props"
      Height          =   435
      Left            =   210
      TabIndex        =   0
      Top             =   2460
      Width           =   1635
   End
   Begin VB.Label lblType 
      AutoSize        =   -1  'True
      Caption         =   "Type"
      Height          =   195
      Left            =   4905
      TabIndex        =   11
      Top             =   615
      Width           =   360
   End
   Begin VB.Label lblDefinedSize 
      AutoSize        =   -1  'True
      Caption         =   "DefinedSize"
      Height          =   195
      Left            =   3345
      TabIndex        =   8
      Top             =   615
      Width           =   855
   End
   Begin VB.Label lblActualSize 
      AutoSize        =   -1  'True
      Caption         =   "ActualSize"
      Height          =   195
      Left            =   1800
      TabIndex        =   6
      Top             =   615
      Width           =   750
   End
   Begin VB.Label lblFieldName 
      AutoSize        =   -1  'True
      Caption         =   "FieldName"
      Height          =   195
      Left            =   210
      TabIndex        =   4
      Top             =   615
      Width           =   750
   End
   Begin VB.Label lblFields 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Left            =   720
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
Dim AdoFileds As Fields


Private Sub cmdExit_Click()
    Unload frmmain ' we unload here
End Sub

Private Sub Command1_Click()
Dim nErr As Error
Dim StrErr As String
Dim TableName As String
Dim m_Filed As Field

On Error GoTo AdoError:
' Example that shows you:
' number of fields in a recordset
' lists the fields
' and also shows some fields properties

    FieldCnt = -1 ' reset our field counter
    ' Clear listboxes
    lstFieldname.Clear
    LstActualSize.Clear
    lstDefinedSize.Clear
    lstType.Clear
    
    TableName = "users" ' Table in the database to open
    Set AdoConn = New Connection
    Set AdoRecSet = New Recordset

    AdoConn.ConnectionString = "Data Source=..\db1.mdb" ' create connection to the database
    AdoConn.Mode = adModeRead ' Open the database for reading only
    AdoConn.Provider = "Microsoft.Jet.OLEDB.4.0" ' set the Provider
    AdoConn.Open ' open the connection
    AdoRecSet.Open TableName, AdoConn, adOpenForwardOnly, adLockReadOnly ' Open Record set
    
    For Each m_Filed In AdoRecSet.Fields ' Loop Though each Filed in the recordset
        ' add the FieldNames to a listbox
        lstFieldname.AddItem m_Filed.Name
        ' Add the ActualSize of the field to a listbox
        LstActualSize.AddItem m_Filed.ActualSize
        ' Add the DefinedSize of the field to a listbox
        lstDefinedSize.AddItem m_Filed.DefinedSize
        ' Add the Type of the field to a listbox see DataTypeEnum for more info on types
        lstType.AddItem m_Filed.Type
    Next
    
    lblFields.Caption = AdoRecSet.Fields.Count ' show the field count
    
    AdoRecSet.Close ' close the recored set always do this first or you have an error
    AdoConn.Close ' close the connection
    ' Destroy the ado objects
    Set AdoConn = Nothing
    Set AdoRecSet = Nothing
    
    Exit Sub
    
AdoError:

    ' This will be triggered if any errors have been found
    For Each nErr In AdoConn.Errors
        StrErr = "Errors: " & AdoConn.Errors.Count & " found" & vbCrLf
        StrErr = StrErr & "Number: " & nErr.Number & vbCrLf
        StrErr = StrErr & "Description: " & nErr.Description & vbCrLf
        StrErr = StrErr & "NativeError: " & nErr.NativeError & vbCrLf
        StrErr = StrErr & "SQLState: " & nErr.SQLState & vbCrLf
        StrErr = StrErr & "Source: " & nErr.Source & vbCrLf
    Next
    
    MsgBox StrErr, vbInformation, "Ado Error" ' show the error
    StrErr = ""
    Unload frmmain
End Sub

