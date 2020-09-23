VERSION 5.00
Begin VB.Form frmmain 
   Caption         =   "Records"
   ClientHeight    =   5100
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7965
   LinkTopic       =   "Form1"
   ScaleHeight     =   5100
   ScaleWidth      =   7965
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOpendb 
      Caption         =   "Show Records"
      Height          =   360
      Left            =   300
      TabIndex        =   2
      Top             =   1920
      Width           =   1695
   End
   Begin VB.ComboBox cboTables 
      Height          =   315
      Left            =   180
      TabIndex        =   0
      Top             =   510
      Width           =   2100
   End
   Begin VB.Label lblTables 
      Caption         =   "Tables:"
      Height          =   255
      Left            =   180
      TabIndex        =   1
      Top             =   225
      Width           =   1020
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' This example shows how to list the tables in a database

Dim AdoConn As Connection

Private Function InitializeDb(lpFileName As String) As Boolean
On Error GoTo AdoError:

    'Load the needed ado objects
    Set AdoConn = New Connection
    Set AdoRecSet = New Recordset
    
    AdoConn.ConnectionString = "Data Source=" & lpFileName ' create connection to the database
    AdoConn.Mode = adModeReadWrite ' Set the mode so we can both read and write to the database
    AdoConn.Provider = "Microsoft.Jet.OLEDB.4.0" ' set the Provider
    AdoConn.Open ' open the connection
    
    InitializeDb = True
    Exit Function
    
AdoError:
    InitializeDb = False
    
End Function

Sub ShowAdoError()
Dim nErr As Error
Dim StrErr As String
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
    
End Sub

Private Sub cmdOpendb_Click()

Dim lpFileName As String
Dim db_Cat As ADOX.Catalog
Dim tb As ADOX.Table
Dim Provider As String
Dim Data_Src As String
Dim db_ok As Boolean

    lpFileName = "..\Tables.mdb" ' Open the Tables test database
    cboTables.Clear ' Clear combo box
    
    db_ok = InitializeDb(lpFileName)
    
    If Not db_ok Then
        ShowAdoError
        Exit Sub
    Else
        Set db_Cat = New ADOX.Catalog 'Create Catalog object
        Set db_Cat.ActiveConnection = AdoConn ' set db_Cat connection to AdoConn connection
        
        
        For Each Table In db_Cat.Tables ' Loop though the Catalog and get all the tables
            ' Fill the Combo box with all the tables in the database
            If Table.Type = "TABLE" Then
                cboTables.AddItem Table.Name
            End If
        Next
        
        lblTables.Caption = "Tables: " & cboTables ' Show the table count

    End If
    
    
   ' Set db_Cat = New ADOX.Catalog ' Create the Catalog to create the database

    Exit Sub
    
CreateERR:
    ' Database was not created
    If Err Then
        MsgBox "Error: " & vbCrLf & "Number: " & Err.Number & vbCrLf _
        & "Source: " & Err.Source & vbCrLf & "Description: " & Err.Description _
        & vbCrLf & "Error Found at: Sub ListUsers", vbInformation
    End If
    
    CreateDataBase = False

    
End Sub

