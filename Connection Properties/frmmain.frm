VERSION 5.00
Begin VB.Form frmmain 
   Caption         =   "Connection Properties"
   ClientHeight    =   3750
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5490
   LinkTopic       =   "Form1"
   ScaleHeight     =   3750
   ScaleWidth      =   5490
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtVal 
      Height          =   495
      Left            =   330
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   2415
      Width           =   4815
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   435
      Left            =   2070
      TabIndex        =   2
      Top             =   3105
      Width           =   1635
   End
   Begin VB.ListBox lstprops 
      Height          =   1620
      Left            =   330
      TabIndex        =   1
      Top             =   450
      Width           =   4755
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show Properties"
      Height          =   435
      Left            =   330
      TabIndex        =   0
      Top             =   3105
      Width           =   1635
   End
   Begin VB.Label lblVal 
      AutoSize        =   -1  'True
      Caption         =   "Value:"
      Height          =   195
      Left            =   330
      TabIndex        =   4
      Top             =   2160
      Width           =   450
   End
   Begin VB.Label lblname 
      AutoSize        =   -1  'True
      Caption         =   "Name:"
      Height          =   195
      Left            =   330
      TabIndex        =   3
      Top             =   180
      Width           =   465
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim AdoConn As Connection
Dim AdoConPropValue() As String ' Array used to hold the values of the ado Properties

Private Sub cmdExit_Click()
    Erase AdoConPropValue
    Unload frmmain
End Sub

Private Sub Command1_Click()
Dim nErr As Error
Dim StrErr As String
On Error GoTo AdoError:
    
    ' This example just shows you thos Properties of a databaes connection
    Me.MousePointer = vbHourglass
    Erase AdoConPropValue()
    lstprops.Clear
    ReDim AdoConPropValue(0)
    Set AdoConn = New Connection ' create ado object
    AdoConn.ConnectionString = "Data Source=..\db1.mdb" ' create connection to the database
    AdoConn.Mode = adModeRead ' Read mode only
    AdoConn.Provider = "Microsoft.Jet.OLEDB.4.0" ' set the Provider
    AdoConn.Open ' open the connection
  
    For Each Item In AdoConn.Properties
        lstprops.AddItem Item.Name
        ' We store the value of the AdoConn.Properties in an array
        ReDim Preserve AdoConPropValue(0 To UBound(AdoConPropValue) + 1) ' Resize array
        AdoConPropValue(UBound(AdoConPropValue)) = Item.Value ' Add value to the array
    Next
    
    AdoConn.Close ' we can now close the connection
    Set AdoConn = Nothing ' destroy the ado object
    Me.MousePointer = vbDefault
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

Private Sub lstprops_Click()
    txtVal.Text = AdoConPropValue(lstprops.ListIndex + 1)
End Sub
