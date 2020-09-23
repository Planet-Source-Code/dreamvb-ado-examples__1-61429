VERSION 5.00
Begin VB.Form frmmain 
   Caption         =   "Error Test"
   ClientHeight    =   990
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2385
   LinkTopic       =   "Form1"
   ScaleHeight     =   990
   ScaleWidth      =   2385
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Make Error"
      Height          =   435
      Left            =   345
      TabIndex        =   0
      Top             =   285
      Width           =   1635
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim AdoConn As Connection

Private Sub Command1_Click()
Dim nErr As Error
Dim StrErr As String
On Error GoTo AdoError:
    
    ' This example to show an error
    'Note that "Data Source=+..\db1.mdb" is spelled incorrectly as it needs no + in it
    
    Set AdoConn = New Connection ' create ado object
    AdoConn.ConnectionString = "Data Source=+..\db1.mdb" ' create connection to the database
    AdoConn.Mode = adModeRead ' Read mode only
    AdoConn.Provider = "Microsoft.Jet.OLEDB.4.0" ' set the Provider
    AdoConn.Open ' open the connection
    
    MsgBox "Connection was made " & "Database version " & AdoConn.Version, vbInformation
    AdoConn.Close ' we can now close the connection
    Set AdoConn = Nothing ' destroy the ado object
    Unload frmmain
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
    
    Set AdoConn = Nothing ' destroy the ado object
    StrErr = ""
    Unload frmmain
    
End Sub

