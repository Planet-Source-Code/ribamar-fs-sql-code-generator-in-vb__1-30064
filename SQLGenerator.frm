VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form SQLGenerator 
   Caption         =   "SQL Generator"
   ClientHeight    =   7005
   ClientLeft      =   195
   ClientTop       =   1035
   ClientWidth     =   10950
   LinkTopic       =   "Form1"
   ScaleHeight     =   7005
   ScaleWidth      =   10950
   Begin VB.CommandButton cmdCopy 
      Caption         =   "&Copy to Clipboard"
      Enabled         =   0   'False
      Height          =   375
      Left            =   9360
      TabIndex        =   7
      Top             =   2880
      Width           =   1455
   End
   Begin VB.TextBox txtValue 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1080
      TabIndex        =   5
      Top             =   2880
      Width           =   3615
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "&Generate"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5400
      TabIndex        =   6
      Top             =   2880
      Width           =   1095
   End
   Begin VB.ListBox lstDataTypes 
      Enabled         =   0   'False
      Height          =   1620
      ItemData        =   "SQLGenerator.frx":0000
      Left            =   9480
      List            =   "SQLGenerator.frx":0028
      TabIndex        =   4
      Top             =   1080
      Width           =   1335
   End
   Begin VB.ListBox lstType 
      Enabled         =   0   'False
      Height          =   1620
      ItemData        =   "SQLGenerator.frx":008D
      Left            =   3240
      List            =   "SQLGenerator.frx":00DC
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   1080
      Width           =   3015
   End
   Begin VB.ListBox lstFields 
      Enabled         =   0   'False
      Height          =   1620
      ItemData        =   "SQLGenerator.frx":01BD
      Left            =   6360
      List            =   "SQLGenerator.frx":01BF
      TabIndex        =   3
      Top             =   1080
      Width           =   3015
   End
   Begin VB.TextBox txtSQL 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Text            =   "SQLGenerator.frx":01C1
      Top             =   3600
      Width           =   10695
   End
   Begin VB.ListBox lstTables 
      Height          =   1620
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   3015
   End
   Begin MSComDlg.CommonDialog dlgSelect 
      Left            =   10320
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Select Database"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label12 
      Caption         =   "Data Type"
      Height          =   255
      Left            =   9720
      TabIndex        =   21
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label11 
      Caption         =   "Select a Field"
      Height          =   255
      Left            =   6720
      TabIndex        =   20
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label10 
      Caption         =   "Select a Table"
      Height          =   255
      Left            =   480
      TabIndex        =   19
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label9 
      BackColor       =   &H80000014&
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9000
      TabIndex        =   18
      Top             =   2880
      Width           =   135
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000014&
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   2880
      Width           =   135
   End
   Begin VB.Label Label7 
      Caption         =   "Value"
      Height          =   255
      Left            =   480
      TabIndex        =   16
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000014&
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5040
      TabIndex        =   15
      Top             =   2880
      Width           =   135
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000014&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9480
      TabIndex        =   14
      Top             =   720
      Width           =   135
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000014&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6360
      TabIndex        =   13
      Top             =   720
      Width           =   135
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000014&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   720
      Width           =   135
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000014&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   11
      Top             =   720
      Width           =   135
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000014&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   135
   End
   Begin VB.Label lblTables 
      Caption         =   "Select a Query Type"
      Height          =   255
      Left            =   3480
      TabIndex        =   9
      Top             =   720
      Width           =   1575
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
   End
   Begin VB.Menu mnuAutor 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "SQLGenerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Code by Ribamar FS ribafs@yahoo.com - http://ribafs.hp10.com.br
'This is a open source program. You is free to change and implementation.
'Thank you for preserv this lines
'Thank your for your feedback of bugs and implementations.

Option Explicit

Dim strType As String
Dim strTableName As String
Dim intCount As Integer
Dim intCountFields As Integer
Dim strDbPath As String
Dim strField2(999) As String
Dim strField1(999) As String
Dim strField As String
Dim strFieldInNew As String
Dim strFieldInCur As String
Dim strFieldView As String
Dim strFieldSelected As String
Dim strSQL As String
Dim strIndex As String
Dim x As Integer

Private Sub cmdCopy_Click()
    txtSQL.SetFocus
    txtSQL.SelStart = 0
    txtSQL.SelLength = Len(txtSQL)
    MsgBox "Altern to your program and press Ctrl+V to paste code!"
    SendKeys "^c"
End Sub

Private Sub cmdGenerate_Click()
    Dim strUpdInput As String
    
    txtSQL = Empty: cmdCopy.Enabled = True
    
    Select Case lstType.Text
        Case "Insert Into"
            txtSQL = strSQL & "INSERT INTO " & lstTables.Text & " (" & strFieldInCur & ") VALUES (" & strFieldInNew & ")" & """"
            
        Case "Select"
            If txtValue.Text = "" Then
                
                txtSQL = strSQL & "SELECT " & strFieldInCur & " FROM " & lstTables.Text & """"
            Else
                If IsNumeric(txtValue.Text) Then
                    txtSQL = strSQL & "SELECT " & strFieldInCur & " FROM " & lstTables.Text & " WHERE " & lstFields.Text & " = " & txtValue & """"
                Else
                    txtSQL = strSQL & "SELECT " & strFieldInCur & " FROM " & lstTables.Text & " WHERE " & lstFields.Text & " = " & "'" & txtValue & "'" & """"
                End If
            End If
            
        Case "Delete"
            If lstDataTypes.Text = "TEXT" Or lstDataTypes.Text = "LONGTEXT" Or lstDataTypes.Text = "DATETIME" Then
                txtSQL = strSQL & "DELETE FROM " & lstTables.Text & " WHERE " & lstFields.Text & " = '" & txtValue & "'" & """"
            Else
                txtSQL = strSQL & "DELETE FROM " & lstTables.Text & " WHERE " & lstFields.Text & " = " & txtValue & """"
            End If

        Case "Update"
            strFieldInCur = InputBox("Current value of " & lstFields.Text & "?")
            strFieldInNew = InputBox("New value to " & lstFields.Text & "?")
            
            If Not IsNumeric(strFieldInNew) Then
                txtSQL = strSQL & "UPDATE " & lstTables.Text & " SET " & lstFields.Text & " = " & "'" & strFieldInNew & "'" & " WHERE " & lstFields.Text & " = " & "'" & strFieldInCur & "'" & """"
            Else
                txtSQL = strSQL & "UPDATE " & lstTables.Text & " SET " & lstFields.Text & " = " & strFieldInNew & " WHERE " & lstFields.Text & " = " & strFieldInCur & """"
            End If
            
        Case "Add Field"
            strFieldInNew = InputBox("Enter Field Name to ADD")
            txtSQL = strSQL & "ALTER TABLE " & lstTables.Text & " ADD COLUMN " & strFieldInNew & " " & lstDataTypes.Text & """"
            
        Case "Alter Field"
            strFieldInNew = InputBox("Enter New Field Name to ALTER")
            txtSQL = strSQL & "ALTER TABLE " & lstTables.Text & " DROP COLUMN " & lstFields.Text & """" & vbCrLf
            txtSQL = txtSQL & "d.Execute ""ALTER TABLE " & lstTables.Text & " ADD COLUMN " & strFieldInNew & " " & lstDataTypes.Text & """"
            
        Case "Drop Table"
                txtSQL = strSQL & "DROP TABLE " & lstTables.Text & """"

        Case "Drop Field"
            txtSQL = strSQL & "ALTER TABLE " & lstTables.Text & " DROP COLUMN " & strFieldInNew & " " & lstFields.Text & """"
        
        Case "Create Index"
            strIndex = InputBox("Enter a Index Name to Create")
            txtSQL = strSQL & "CREATE INDEX " & strIndex & " ON " & lstTables.Text & " (" & strFieldInCur & ")" & """"
            
        Case "Create IndexUnique"
            strIndex = InputBox("Enter a Index Unique Name to Create a Index WITH DISALLOW NULL")
            txtSQL = strSQL & "CREATE UNIQUE INDEX " & strIndex & " ON " & lstTables.Text & " (" & lstFields.Text & ")" & " WITH DISALLOW NULL" & """"
            
        Case "Drop Index"
            txtSQL = strSQL & "DROP INDEX " & lstFields.Text & " ON " & lstTables.Text & """"
            
        Case "Avg"
            strFieldInNew = InputBox("Enter a Alias to Avg from " & lstFields.Text)
            txtSQL = strSQL & "SELECT AVG" & "(" & lstFields.Text & ")" & " AS (" & strFieldInNew & ") FROM " & lstTables.Text & " WHERE " & lstFields.Text & "=" & txtValue & """"
            
        Case "Count"
            strFieldInCur = InputBox("Enter a Field Name to Count")
            strFieldInNew = InputBox("Enter a Alias to Count from " & strFieldInCur)
            txtSQL = strSQL & "SELECT " & lstFields.Text & " COUNT (" & strFieldInCur & ") FROM " & lstTables.Text & " AS " & "(" & strFieldInNew & ")" & """"
            
        Case "Min"
            strFieldInNew = InputBox("Enter a Alias to Min from " & lstFields.Text)
            txtSQL = strSQL & "SELECT MIN" & "(" & lstFields.Text & ")" & " AS (" & strFieldInNew & ") FROM " & lstTables.Text & """"
            
        Case "Max"
            strFieldInNew = InputBox("Enter a Alias to Max from " & lstFields.Text)
            txtSQL = strSQL & "SELECT MAX" & "(" & lstFields.Text & ")" & " AS (" & strFieldInNew & ") FROM " & lstTables.Text & """"
            
        Case "StDev"
            strFieldInNew = InputBox("Enter a Alias to StDev from " & lstFields.Text)
            txtSQL = strSQL & "SELECT STDEV" & "(" & lstFields.Text & ")" & " AS (" & strFieldInNew & ") FROM " & lstTables.Text & """"
            
        Case "StDevP"
            strFieldInNew = InputBox("Enter a Alias to StDevP from " & lstFields.Text)
            txtSQL = strSQL & "SELECT STDEVP" & "(" & lstFields.Text & ")" & " AS (" & strFieldInNew & ") FROM " & lstTables.Text & """"
            
        Case "Sum"
            strFieldInNew = InputBox("Enter a Alias to Sum from " & lstFields.Text)
            txtSQL = strSQL & "SELECT SUM" & "(" & lstFields.Text & ")" & " AS (" & strFieldInNew & ") FROM " & lstTables.Text & """"
            
        Case "Var"
            strFieldInNew = InputBox("Enter a Alias to Var from " & lstFields.Text)
            txtSQL = strSQL & "SELECT VAR" & "(" & lstFields.Text & ")" & " AS (" & strFieldInNew & ") FROM " & lstTables.Text & """"
            
        Case "VarP"
            strFieldInNew = InputBox("Enter a Alias to VarP from " & lstFields.Text)
            txtSQL = strSQL & "SELECT VARP" & "(" & lstFields.Text & ")" & " AS (" & strFieldInNew & ") FROM " & lstTables.Text & """"
            
        Case "Group By"
            strFieldInNew = InputBox("Enter a Field Name to Group By")
            txtSQL = strSQL & "SELECT " & strFieldInCur & " FROM " & lstTables.Text & " GROUP BY " & strFieldInNew & """"
            
        Case "Having"
            strFieldInNew = InputBox("Enter a Field Name to Group By")
            strFieldInCur = InputBox("Enter a Field Name to Having")
            txtSQL = strSQL & "SELECT " & strFieldInCur & " FROM " & lstTables.Text & " GROUP BY " & strFieldInNew & " HAVING " & strFieldInCur & "=" & txtValue.Text & """"
            
        Case "Order By"
            strFieldInNew = InputBox("Enter a Field Name to Order By")
            strFieldInCur = InputBox("Enter ASC to Ascendent or DESC to Descentend")
            txtSQL = strSQL & "SELECT " & strFieldInCur & " FROM " & lstTables.Text & " ORDER BY " & strFieldInNew & " " & strFieldInCur & """"
            
        Case "Top"
            strFieldInCur = InputBox("Enter ASC to Ascendent or DESC to Descentend")
            txtSQL = strSQL & "SELECT TOP " & txtValue.Text & " FROM " & lstTables.Text & " ORDER BY " & lstFields.Text & " " & strFieldInCur & """"
        
        Case "In"
            strFieldInNew = InputBox("Enter a String to Search In. Samples: 'Anthony', 'A*', 6, 203")
            If IsNumeric(strFieldInNew) Then
                txtSQL = strSQL & "SELECT " & lstFields.Text & " FROM " & lstTables.Text & " IN " & "'" & strDbPath & "'" & " WHERE " & lstFields.Text & " LIKE " & strFieldInNew & """"
            Else
                txtSQL = strSQL & "SELECT " & lstFields.Text & " FROM " & lstTables.Text & " IN " & "'" & strDbPath & "'" & " WHERE " & lstFields.Text & " LIKE " & "'" & strFieldInNew & "'" & """"
            End If
            
        Case Else
    End Select
End Sub

Private Sub lstType_Click()
    strFieldInNew = "": strFieldInCur = "": strFieldInNew = ""
    strFieldSelected = "": x = 0
    
    Select Case lstType.Text
        Case "Update", "Alter Field"
            lstDataTypes.Enabled = True: lstFields.Enabled = True
        Case "Avg", "Count", "Min", "Max", "Insert Into", "Select", "Delete", "Drop Field", _
            "Create Index", "Create IndexUnique", "Drop Index", "StDev", "StDevP", "Sum", "Var", _
            "VarP", "Group By", "Having", "Order By", "Top", "In"
            lstFields.Enabled = True: txtValue.Enabled = True
        Case "Add Field"
            lstDataTypes.Enabled = True: lstFields.Enabled = False
        Case Else
            lstFields.Enabled = False: lstDataTypes.Enabled = False
    End Select

End Sub

Private Sub cmdSelect_Click()
    On Error GoTo errorhandler

    Dim d As DAO.Database
    Dim td As DAO.TableDef
    Dim matTable(999) As String
            
    dlgSelect.CancelError = True
    dlgSelect.Filter = "Access Database (*.mdb)|*.mdb"
    dlgSelect.ShowOpen
    cmdSelect.Enabled = False
    DoEvents
    
    strDbPath = dlgSelect.FileName
    Set d = OpenDatabase(strDbPath)
    lstTables.Clear
    
    For Each td In d.TableDefs
        If Mid(td.Name, 1, 4) <> "MSys" Then
            intCount = intCount + 1
            matTable(intCount) = td.Name
            lstTables.AddItem (matTable(intCount))
        End If
    Next
    cmdSelect.Enabled = True

    d.Close
    Set d = Nothing
    Set td = Nothing
errorhandler:

End Sub

Private Sub Form_Load()
    strSQL = "'" & Space(60) & "SQL GENERATOR" & vbCrLf & vbCrLf & vbCrLf
    strSQL = strSQL & "'Use suggestion:" & vbCrLf
    strSQL = strSQL & "'In VB and Access 2000 make reference to DAO:" & vbCrLf
    strSQL = strSQL & "'In General Declaration:" & vbCrLf & vbCrLf
    strSQL = strSQL & "Dim d As Database" & vbCrLf
    strSQL = strSQL & "'In Event Form_Load():" & vbCrLf & vbCrLf
    strSQL = strSQL & "Set d = OpenDatabase(strPathDatabase)" & vbCrLf
    strSQL = strSQL & "'In Procedure to use SQL:" & vbCrLf & vbCrLf
    strSQL = strSQL & "d.Execute """
End Sub

Private Sub lstFields_Click()
    'If lstType.Text = "Insert Into" Then
        x = x + 1
    
        If x = 1 Then
            strFieldInCur = strFieldInCur & lstFields.Text
        Else
            strFieldInCur = strFieldInCur & ", " & lstFields.Text
        End If
        
    If lstType.Text = "Insert Into" Then
        strFieldSelected = InputBox("Enter value of " & lstFields.Text)
        If Not IsNumeric(lstFields.Text) Then
            If x <= 1 Then
                strFieldInNew = strFieldInNew & "'" & strFieldSelected & "'"
            Else
                strFieldInNew = strFieldInNew & ", " & "'" & strFieldSelected & "'"
            End If
        Else
            If x <= 1 Then
                strFieldInNew = strFieldInNew & strFieldSelected
            Else
                strFieldInNew = strFieldInNew & ", " & strFieldSelected
            End If
        End If
    End If
End Sub

Private Sub lstTables_Click()
    Dim d As Database
    Dim td As TableDef
    Dim f As DAO.Field
    Dim matField(99) As String
    
    cmdGenerate.Enabled = True: lstType.Enabled = True
    strTableName = lstTables.Text
    lstFields.Clear
    Set d = OpenDatabase(strDbPath)
    Set td = d.TableDefs(strTableName)
    
    intCount = 0: strField = "": strFieldInNew = "": strFieldInCur = ""
    
    For Each f In td.Fields
        intCount = intCount + 1
        matField(intCount) = f.Name
        lstFields.AddItem (matField(intCount))
        strField2(intCount) = matField(intCount)
        
        If intCount = 1 And lstType.Text = "Insert Into" Then
            strField = strField & matField(intCount)
        Else
            strField = strField & ", " & matField(intCount)
        End If
        
    Next f

    d.Close
    Set td = Nothing
    Set d = Nothing
End Sub

Private Sub mnuAutor_Click()
    Load frmAbout
    frmAbout.Show
End Sub

Private Sub mnuHelp_Click()
    Shell "c:\windows\notepad.exe " & App.Path & ("\help.txt"), vbMaximizedFocus
End Sub

