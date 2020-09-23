VERSION 5.00
Begin VB.Form frmmain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DataEntry"
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8505
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   8505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   405
      Left            =   5790
      TabIndex        =   21
      Top             =   4365
      Width           =   1215
   End
   Begin VB.CommandButton cmdSaveUpdate 
      Caption         =   "Save/Update"
      Enabled         =   0   'False
      Height          =   405
      Left            =   2865
      TabIndex        =   20
      Top             =   4365
      Width           =   1365
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Height          =   405
      Left            =   1455
      TabIndex        =   19
      Top             =   4365
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add New"
      Height          =   405
      Left            =   120
      TabIndex        =   18
      Top             =   4365
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   405
      Left            =   4380
      TabIndex        =   17
      Top             =   4365
      Width           =   1215
   End
   Begin VB.TextBox txtDOB 
      Height          =   330
      Left            =   4290
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   3375
      Width           =   1455
   End
   Begin VB.TextBox txtAge 
      Height          =   330
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   2850
      Width           =   915
   End
   Begin VB.ComboBox cboCounty 
      Height          =   315
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   2340
      Width           =   2010
   End
   Begin VB.TextBox txtAddr 
      Height          =   855
      Left            =   4320
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Top             =   1290
      Width           =   3555
   End
   Begin VB.TextBox txtLastName 
      Height          =   330
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   840
      Width           =   3540
   End
   Begin VB.TextBox txtFirstName 
      Height          =   330
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   360
      Width           =   2880
   End
   Begin VB.Frame frmGender 
      Caption         =   "Gender"
      Height          =   930
      Left            =   6180
      TabIndex        =   8
      Top             =   2835
      Width           =   1950
      Begin VB.OptionButton OptGender 
         Caption         =   "Female"
         Height          =   195
         Index           =   1
         Left            =   165
         TabIndex        =   16
         Top             =   570
         Width           =   1605
      End
      Begin VB.OptionButton OptGender 
         Caption         =   "Male"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   15
         Top             =   270
         Width           =   1140
      End
   End
   Begin VB.ListBox lstUsers 
      Height          =   3570
      Left            =   150
      TabIndex        =   0
      Top             =   465
      Width           =   1935
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   1
      X1              =   165
      X2              =   8355
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   165
      X2              =   8355
      Y1              =   4215
      Y2              =   4215
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "Date of Birth"
      Height          =   195
      Index           =   5
      Left            =   2415
      TabIndex        =   7
      Top             =   3450
      Width           =   1290
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "County:"
      Height          =   195
      Index           =   4
      Left            =   2430
      TabIndex        =   6
      Top             =   2490
      Width           =   1290
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "Address:"
      Height          =   195
      Index           =   3
      Left            =   2430
      TabIndex        =   5
      Top             =   1380
      Width           =   1290
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "Age"
      Height          =   195
      Index           =   2
      Left            =   2430
      TabIndex        =   4
      Top             =   2940
      Width           =   1290
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "Last Name"
      Height          =   195
      Index           =   1
      Left            =   2430
      TabIndex        =   3
      Top             =   900
      Width           =   1290
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "First Name:"
      Height          =   195
      Index           =   0
      Left            =   2430
      TabIndex        =   2
      Top             =   420
      Width           =   1290
   End
   Begin VB.Label lblUsers 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Users"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   150
      TabIndex        =   1
      Top             =   165
      Width           =   495
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' OK now this example will show you how to
' Load in the database
' Load in two tables
' View the data from each table
' Navigating though data
' Preforming a simple Queries
' Adding new records
' Deleteing Records
' Editing, Updateing Exsiting Records

' I have used a new database it is the same as in last exmples
' but this time has some more fields and a new table

Enum EditOption
    Add = 0
    Edit
End Enum

Dim m_def_EditMode As EditOption
Dim AdoConn As Connection
Dim AdoRecSet As Recordset
Dim Table_Name As String ' this is the table in the db to load
Dim m_Record_ID As Long ' Current ID of record

Private Sub LockControlsA(aLock As Boolean)
Dim I As Integer
    For I = 0 To frmmain.Controls.Count - 1
        Select Case TypeName(frmmain.Controls(I))
            Case "TextBox", "ComboBox": frmmain.Controls(I).Locked = aLock
        End Select
    Next
End Sub

Private Sub ResetFormFields()
Dim I As Integer
    OptGender(0).Value = True
    
    ' Clear all controls of there values
    For I = 0 To frmmain.Controls.Count - 1
        Select Case TypeName(frmmain.Controls(I))
            Case "TextBox": frmmain.Controls(I).Text = "": frmmain.Controls(I).Locked = False
            Case "ComboBox": cboCounty.Text = "": frmmain.Controls(I).Locked = False
        End Select
    Next
End Sub

Private Function FixFieldData(lpData As String) As Variant
    If Len(lpData) = 0 Then
        FixFieldData = "Null"
        Exit Function
    Else
        FixFieldData = lpData
    End If
End Function

Private Sub DoSaveData(lpAction As EditOption)
Dim MySQL As String

On Error GoTo OpenErr:


    If lpAction = Add Then ' user requested to add a new Record
        If AdoRecSet.State = adStateOpen Then AdoRecSet.Close ' if RecordSet is open we first must close it
        AdoRecSet.Open Table_Name, AdoConn, adOpenKeyset, adLockPessimistic

        AdoRecSet.AddNew    ' Add New Record
        Call SaveFileds     ' Save the Fields
        AdoRecSet.Update    ' Update the database
        AdoRecSet.Close
        MsgBox "The new record has now been added", vbInformation
        LoadRecordSet ' Load in the users table
        ListUsers lstUsers ' Show the users in the list
        Exit Sub
    Else ' user requested to edit a Record
        MySQL = "SELECT ID,FirstName,LastName,Gender,Age,Address,DOB,County FROM " _
        & Table_Name & " WHERE ID =" & m_Record_ID & ""
        
        If AdoRecSet.State = adStateOpen Then AdoRecSet.Close ' if RecordSet is open we first must close it
        AdoRecSet.Open MySQL, AdoConn, adOpenKeyset, adLockPessimistic
        Call SaveFileds
        AdoRecSet.Update
        AdoRecSet.Close
        LoadRecordSet ' Load in the users table
        ListUsers lstUsers ' Show the users in the list
        MsgBox "The new record has now been updated", vbInformation
    End If
    
OpenErr:
    ' Show a messagebox if we found any errors
    If Err Then
        MsgBox "Error: " & vbCrLf & "Number: " & Err.Number & vbCrLf _
        & "Source: " & Err.Source & vbCrLf & "Description: " & Err.Description _
        & vbCrLf & "Error Found at: Sub DoSaveData", vbInformation
    End If
    
End Sub

Private Sub SaveFileds()
    On Error Resume Next
    With AdoRecSet
        !FirstName = FixFieldData(txtFirstName.Text)
        !LastName = FixFieldData(txtLastName.Text)
        
        If OptGender(0).Value Then
            !Gender = "M"
        End If
        
        If OptGender(1).Value Then
            !Gender = "F"
        End If
        
        !Age = Val(txtAge.Text)
        !Address = FixFieldData(txtAddr.Text)
        If IsDate(txtDOB.Text) Then !DOB = CDate(txtDOB.Text) Else !DOB = Date
        
        !County = cboCounty.ListIndex + 1
    End With
    
    If Err Then MsgBox Err.Description
    
End Sub

Private Function DeleteRecord() As Boolean
On Error GoTo DeleteErr:
Dim MySQL As String

    MySQL = "DELETE FROM " & Table_Name & " WHERE ID=" & m_Record_ID & "" ' Query used to delete the record
    AdoConn.Execute MySQL ' execute the Query Above
    DeleteRecord = True ' send back good value

    Exit Function
DeleteErr:
    If Err Then Err.Clear: DeleteRecord = False ' return bad value
    
End Function

Private Sub ShowUserInfo(m_byFirstName As String)
Dim MySQL As String

    ' This sub will show information on the selected user
    ' we will recive this information from the database using a simple Querie based on the users First Name
    
    MySQL = "SELECT ID,FirstName,LastName,Gender,Age,Address,DOB,County FROM " _
    & Table_Name & " WHERE FirstName ='" & m_byFirstName & "'"
    ' so what the code in simple terms is really saying is
    ' Select all the information from the users table were there First name is equal to m_byFirstName
    
    AdoRecSet.Open MySQL, AdoConn, adOpenForwardOnly, adLockReadOnly
    
    m_Record_ID = Val(AdoRecSet("ID")) ' store the ID of the record
    ' Fill in textboxes with the persons information
    txtFirstName.Text = AdoRecSet("FirstName") ' FirstName
    txtLastName.Text = AdoRecSet("LastName") ' LastName
    txtAddr.Text = AdoRecSet("Address") ' Address
    cboCounty.ListIndex = Val(AdoRecSet("County") - 1) 'County
    txtAge.Text = AdoRecSet("Age") 'Age
    txtDOB.Text = AdoRecSet("DOB") 'Date of Birth
    
    If UCase(AdoRecSet("Gender")) = "M" Then ' is person male or female
        OptGender(0).Value = True ' Male
    Else
        OptGender(1).Value = True ' Female
    End If
    
    MySQL = ""

    AdoRecSet.Close ' Clsoe the table
    
End Sub

Private Function ListCountys(cbBox As ComboBox)
    ' This places all the County names into a combobox
    Do While Not AdoRecSet.EOF ' Loop until we reach the end of the record
        With AdoRecSet
            cbBox.AddItem AdoRecSet("Name") ' add in the firstName
            .MoveNext ' Move to the next aviable record
        End With
    Loop
    
    AdoRecSet.Close ' Close recored set
End Function

Private Sub ListUsers(lbBox As ListBox)
On Error Resume Next

    lbBox.Clear
    
    ' This places all the users, first names into a listbox
    Do While Not AdoRecSet.EOF ' Loop until we reach the end of the record
        With AdoRecSet
            lbBox.AddItem AdoRecSet("FirstName") ' add in the firstName
            .MoveNext ' Move to the next aviable record
        End With
    Loop
    
    lblUsers.Caption = "Users: " & lstUsers.ListCount  ' display the number of users found
    AdoRecSet.Close ' Close the recored set
    
    If Err Then
        MsgBox "Error: " & vbCrLf & "Number: " & Err.Number & vbCrLf _
        & "Source: " & Err.Source & vbCrLf & "Description: " & Err.Description _
        & vbCrLf & "Error Found at: Sub ListUsers", vbInformation
    End If
    
End Sub

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

Sub CloseAndCleanUp()
    ' we use this sub to clean eveything we have used
    If AdoRecSet.State = adStateOpen Then AdoRecSet.Close ' close the recored set
    If AdoConn.State = adStateOpen Then AdoConn.Close ' close the connection
    
    ' Destroy the ado objects and clean up other variables
    Table_Name = ""
    m_Record_ID = 0
    Set AdoConn = Nothing
    Set AdoRecSet = Nothing
End Sub

Private Function InitializeDb() As Boolean
On Error GoTo AdoError:

    'Load the needed ado objects
    Set AdoConn = New Connection
    Set AdoRecSet = New Recordset
    
    AdoConn.ConnectionString = "Data Source=..\db2.mdb" ' create connection to the database
    AdoConn.Mode = adModeReadWrite ' Set the mode so we can both read and write to the database
    AdoConn.Provider = "Microsoft.Jet.OLEDB.4.0" ' set the Provider
    AdoConn.Open ' open the connection
    
    InitializeDb = True
    Exit Function
    
AdoError:
    InitializeDb = False
    
End Function

Private Function LoadRecordSet() As Boolean
On Error GoTo LoadRecErr:
    
    AdoRecSet.Open Table_Name, AdoConn, adOpenForwardOnly, adLockReadOnly ' Open Record set
    LoadRecordSet = True ' send back good value
    Exit Function ' exit the code block
LoadRecErr:
    ' Report errors found
    MsgBox "Error:" & vbCrLf _
    & "Description: " & Err.Description _
    & vbCrLf & "Source: " & Err.Source, vbInformation, "Function LoadRecordSet()"
    LoadRecordSet = False
    
End Function

Private Sub cmdAdd_Click()
    'Add a new record
    m_def_EditMode = Add
    lstUsers.Enabled = False
    cmdEdit.Enabled = False
    cmdDelete.Caption = "Reset"
    cmdSaveUpdate.Enabled = True
    Call ResetFormFields
    
End Sub

Private Sub cmdDelete_Click()

    If cmdDelete.Caption = "Delete" Then ' Delete mode
        If MsgBox("Are you sure you want to delete this record", vbYesNo Or vbQuestion) = vbNo Then Exit Sub
        
        If Not DeleteRecord Then
            MsgBox "The Record was not deleted", vbCritical
            Exit Sub
        Else
            MsgBox "The Record has been deleted", vbInformation
            LoadRecordSet ' Load in the users table
            ListUsers lstUsers ' Show the users in the list
            If lstUsers.ListCount > 0 Then lstUsers.ListIndex = 0
            If lstUsers.ListCount = 0 Then cmdDelete.Enabled = False
        End If
        Exit Sub
    Else
        ' Reset Mode
        Call ResetFormFields
    End If
    
End Sub

Private Sub cmdEdit_Click()
    ' Edit Record
    m_def_EditMode = Edit
    cmdEdit.Enabled = False
    lstUsers.Enabled = False
    cmdDelete.Caption = "Reset"
    cmdSaveUpdate.Enabled = True
    cmdAdd.Enabled = False
    LockControlsA False
End Sub

Private Sub cmdExit_Click()
    CloseAndCleanUp
    Unload frmmain
End Sub

Private Sub cmdSaveUpdate_Click()
    DoSaveData m_def_EditMode
    cmdDelete.Caption = "Delete"
    cmdEdit.Enabled = True
    lstUsers.Enabled = True
    
    If m_def_EditMode = Add Then
        lstUsers.ListIndex = lstUsers.ListCount - 1
    End If
    
    LockControlsA True
End Sub



Private Sub Form_Load()
Dim InitGood As Boolean
    
    Table_Name = "tbUsers" ' table to load
    ' First thing to do is to Initialize the ado object and open the database
    InitGood = InitializeDb()
    
    If Not InitGood Then ' check for any errors
        ShowAdoError ' show error message
        Unload frmmain 'unload the form
        Exit Sub
    End If
   
    If Not LoadRecordSet() Then 'next we load the recordset
        Unload frmmain
        Exit Sub
    Else
        'Fill the listbox with all the users first name
        ListUsers lstUsers
    End If
    
    Table_Name = "tbCounty"
    ' Next we load in the County table to fill the Countys combo box
    If Not LoadRecordSet Then
        MsgBox "There was an error loading the " & Table_Name & " Table.", vbInformation
        CloseAndCleanUp ' Do cleanup
        Unload frmmain ' Unload the form
        Exit Sub
    Else
        ListCountys cboCounty ' Add the Countys to the combobox
        Table_Name = "tbUsers" ' change back to users table
    End If
    
    If lstUsers.ListCount > 0 Then lstUsers.ListIndex = 0
    
End Sub

Private Sub lstUsers_Click()
    ShowUserInfo lstUsers.Text
End Sub
