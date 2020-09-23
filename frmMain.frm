VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sql Server Compare"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9855
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   9855
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picSplitterIcon 
      Height          =   375
      Left            =   0
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":0442
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   23
      Top             =   5040
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picBG 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4215
      Left            =   4680
      ScaleHeight     =   4215
      ScaleWidth      =   5175
      TabIndex        =   18
      Top             =   0
      Width           =   5175
      Begin VB.PictureBox picResizer 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   100
         Left            =   0
         ScaleHeight     =   105
         ScaleWidth      =   5175
         TabIndex        =   22
         Top             =   2040
         Width           =   5175
      End
      Begin VB.PictureBox picSplitter 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   100
         Left            =   0
         MousePointer    =   99  'Custom
         ScaleHeight     =   105
         ScaleWidth      =   5175
         TabIndex        =   21
         Top             =   2040
         Width           =   5175
      End
      Begin RichTextLib.RichTextBox rtbDetails 
         Height          =   2055
         Left            =   0
         TabIndex        =   19
         Top             =   2160
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   3625
         _Version        =   393217
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         OLEDragMode     =   0
         OLEDropMode     =   0
         TextRTF         =   $"frmMain.frx":074C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.TreeView tvwSQL 
         Height          =   2055
         Left            =   0
         TabIndex        =   20
         Top             =   0
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   3625
         _Version        =   393217
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   6
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Select databases to compare...                          "
      Height          =   2055
      Left            =   0
      TabIndex        =   9
      Top             =   2160
      Width           =   4575
      Begin VB.TextBox txtScripting 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         ScrollBars      =   1  'Horizontal
         TabIndex        =   17
         Top             =   1200
         Width           =   4335
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit"
         Height          =   375
         Left            =   1650
         TabIndex        =   16
         Top             =   1560
         Width           =   1335
      End
      Begin VB.CommandButton cmdExport 
         Caption         =   "Export"
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   1560
         Width           =   1335
      End
      Begin VB.CommandButton cmdCompare 
         Caption         =   "Compare"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3120
         TabIndex        =   12
         Top             =   1560
         Width           =   1335
      End
      Begin VB.ComboBox cmbDatabases1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   360
         Width           =   3015
      End
      Begin VB.ComboBox cmbDatabases2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   720
         Width           =   3015
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Working DB"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   315
         Index           =   1
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Client DB"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   315
         Index           =   0
         Left            =   240
         TabIndex        =   13
         Top             =   720
         Width           =   1215
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   4245
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   661
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   17286
            MinWidth        =   17286
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Connection                                                         "
      Height          =   2055
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   4575
      Begin VB.CommandButton cmdSettings 
         Caption         =   "Settings"
         Height          =   375
         Left            =   3120
         TabIndex        =   27
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   1000
         Left            =   120
         TabIndex        =   7
         Top             =   940
         Width           =   2895
         Begin VB.TextBox txtPassword 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   960
            TabIndex        =   2
            Top             =   720
            Width           =   1815
         End
         Begin VB.TextBox txtLogin 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   960
            TabIndex        =   1
            Top             =   360
            Width           =   1815
         End
         Begin VB.TextBox txtServer 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   960
            TabIndex        =   0
            Top             =   0
            Width           =   1815
         End
         Begin VB.Label lblText 
            Caption         =   "Password"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   0
            TabIndex        =   26
            Top             =   720
            Width           =   855
         End
         Begin VB.Label lblText 
            Caption         =   "Username"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   25
            Top             =   360
            Width           =   855
         End
         Begin VB.Label lblText 
            Caption         =   "Server"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   24
            Top             =   0
            Width           =   855
         End
      End
      Begin VB.CommandButton cmdConnect 
         Caption         =   "Connect"
         Height          =   375
         Left            =   3120
         TabIndex        =   6
         Top             =   1560
         Width           =   1335
      End
      Begin VB.OptionButton optSQLVerif 
         Caption         =   "SQL Server Verification"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Value           =   -1  'True
         Width           =   2535
      End
      Begin VB.OptionButton optTrusted 
         Caption         =   "Trusted (NT Verification)"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   2895
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private bSplitStarted As Boolean
Private iSplitterX As Single
Private iCurrentX As Single

Private ScriptingIsProcessed As Boolean

Private Enum ScriptObjects
    enmScriptStoredProcedure = 0
    enmScriptView = 1
    enmScriptTable = 2
    enmScriptAlterTable = 3
End Enum

Private iScriptType, iScript2Type

Private SqlSrv As Object

Private DB1COL As Collection

Private Sub cmdCompare_Click()
    If StrComp(Me.cmbDatabases1.Text, Me.cmbDatabases2.Text, vbTextCompare) = 0 Then
        MsgBox "Compare on same database choosen..", vbExclamation, "Error"
        Exit Sub
    End If
    CompareDB
End Sub

Private Sub cmdConnect_Click()
    ConnectSQL
End Sub

Private Function ConnectSQL() As Boolean
    On Error GoTo Err_Handler
    Screen.MousePointer = vbHourglass
    Set SqlSrv = CreateObject("SQLDMO.SQLServer")
    Me.StatusBar1.SimpleText = "Connecting..please wait"
    SqlSrv.LoginTimeOut = 3
    If Me.optSQLVerif.Value Then
        SqlSrv.Connect Me.txtServer.Text, Me.txtLogin.Text, Me.txtPassword.Text
    ElseIf Me.optTrusted.Value Then
        SqlSrv.LoginSecure = True
        SqlSrv.Connect
    End If
    If SqlSrv.verifyconnection Then
        Me.StatusBar1.SimpleText = "Connection succeeded..Retrieving databases"
        getDatabases
        Me.cmdCompare.Enabled = True
    Else
        MsgBox "Not Connected.. verify settings", vbExclamation, "Error"
    End If
    Me.StatusBar1.SimpleText = ""
    Screen.MousePointer = vbDefault
    On Error GoTo 0
    Exit Function
Err_Handler:
    Me.StatusBar1.SimpleText = ""
    Screen.MousePointer = vbDefault
    MsgBox Err.Description
    On Error GoTo 0
    Err.Clear
End Function

Private Function getDatabases()
    On Error GoTo Err_Handler
    Me.cmbDatabases1.Clear
    Me.cmbDatabases2.Clear
    Dim oObject As Object
    For Each oObject In SqlSrv.Databases
        Me.cmbDatabases1.AddItem oObject.Name
        Me.cmbDatabases2.AddItem oObject.Name
    Next oObject
    If Me.cmbDatabases1.ListCount > 0 Then Me.cmbDatabases1.ListIndex = 0
    If Me.cmbDatabases2.ListCount > 0 Then Me.cmbDatabases2.ListIndex = 0
    On Error GoTo 0
    Exit Function
Err_Handler:
    MsgBox Err.Description
    On Error GoTo 0
    Err.Clear
End Function

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdExport_Click()
    Me.rtbDetails.SaveFile App.Path & "\Changes.txt", 1
End Sub

Private Sub cmdSettings_Click()
    frmSettings.FormDoModal
End Sub

Private Sub Form_Load()
    ScriptSettings.ScriptColumns = True
    ScriptSettings.ScriptOwnerDiff = False
    ScriptSettings.ScriptStoredProcedures = True
    ScriptSettings.ScriptTables = True
    ScriptSettings.ScriptViews = True
    ScriptSettings.ScriptAutoProcess = False
    Me.picResizer.MouseIcon = Me.picSplitterIcon.Picture
    Me.picSplitter.MouseIcon = Me.picSplitterIcon.Picture
    ResizeControlsIR
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not DB1COL Is Nothing Then Set DB1COL = Nothing
    If Not SqlSrv Is Nothing Then
        On Error Resume Next
        If SqlSrv.verifyconnection Then
            SqlSrv.Close
        End If
        Set SqlSrv = Nothing
    End If
End Sub

Private Function WriteQuery(ByVal Q As String) As String
    Dim Fi As Long
    Fi = FreeFile
    Open App.Path & "\UpdateQuery.sql" For Output As Fi
    Print #Fi, "USE [" & Me.cmbDatabases2.Text & "]" & vbCrLf & vbCrLf & "GO" & vbCrLf
    Print #Fi, """" & Q & """"
    Close Fi
    WriteQuery = App.Path & "\UpdateQuery.sql"
End Function

Private Sub optSQLVerif_Click()
    Me.Frame2.Visible = True
    Me.txtServer.SetFocus
End Sub

Private Sub optTrusted_Click()
    Me.Frame2.Visible = False
End Sub

Private Function CompareDB()
    On Error GoTo Err_Handler
    Screen.MousePointer = vbHourglass
    If Not DB1COL Is Nothing Then Set DB1COL = Nothing
    Set DB1COL = New Collection
    
    Me.rtbDetails.Text = ""
    Me.tvwSQL.Nodes.Clear
    Me.txtScripting.Text = ""
    
    KillQueryFile
    
    DoEvents
    
    ScriptingIsProcessed = False
    
    Dim SQLObj As Object
    
    Dim SQLColumn As Object
    
    Dim cPropOrg As clsProps
    
    Dim cSP As clsStorProcs
    Dim cView As clsViews
    Dim cTable As clsTables
    Dim cColumn As clsColumns
    
    Set cPropOrg = New clsProps
    cPropOrg.SQLName = Me.cmbDatabases1.Text
    cPropOrg.SQLOwner = SqlSrv.Databases(Me.cmbDatabases1.Text).Owner
    
    Me.StatusBar1.SimpleText = "Retrieving stored procedures... " & Me.cmbDatabases1.Text
    For Each SQLObj In SqlSrv.Databases(Me.cmbDatabases1.Text).StoredProcedures
        If Not SQLObj.SystemObject Then
            Set cSP = New clsStorProcs
            cSP.SQLSPName = SQLObj.Name
            cSP.SQLSPCommand = SQLObj.Text
            cSP.SQLSPShortCmd = RemoveUnwantedChars(cSP.SQLSPCommand)
            cSP.SQLSPProcessed = False
            cSP.SQLSPOwner = SQLObj.Owner
            cSP.SQLSPID = SQLObj.ID
            cPropOrg.StorProcCol.Add cSP, "SP" & cSP.SQLSPName & cSP.SQLSPOwner
        End If
    Next
    
    Me.StatusBar1.SimpleText = "Retrieving tables... " & Me.cmbDatabases1.Text
    For Each SQLObj In SqlSrv.Databases(Me.cmbDatabases1.Text).Tables
        If Not SQLObj.SystemObject Then
            Set cTable = New clsTables
            cTable.SQLTableFieldsCount = SQLObj.Columns.Count
            cTable.SQLTableName = SQLObj.Name
            cTable.SQLTableProcessed = False
            cTable.SQLTableOwner = SQLObj.Owner
            cTable.SQLTableID = SQLObj.ID
            For Each SQLColumn In SQLObj.Columns
                Set cColumn = New clsColumns
                cColumn.SQLColIdentifier = SQLColumn.Identity
                cColumn.SQLColAllowNULL = SQLColumn.AllowNulls
                cColumn.SQLColName = SQLColumn.Name
                cColumn.SQLColType = SQLColumn.DataType
                cColumn.SQLColID = SQLColumn.ID
                cTable.ColumnCol.Add cColumn, cColumn.SQLColName
            Next
            cPropOrg.TableCol.Add cTable, "TABLE" & cTable.SQLTableName & cTable.SQLTableOwner
        End If
    Next
    
    Me.StatusBar1.SimpleText = "Retrieving views... " & Me.cmbDatabases1.Text
    For Each SQLObj In SqlSrv.Databases(Me.cmbDatabases1.Text).Views
        Set cView = New clsViews
        cView.SQLViewName = SQLObj.Name
        cView.SQLViewCommand = SQLObj.Text
        cView.SQLViewShortCmd = RemoveUnwantedChars(cView.SQLViewCommand)
        cView.SQLViewProcessed = False
        cView.SQLViewOwner = SQLObj.Owner
        cView.SQLViewID = SQLObj.ID
        cPropOrg.ViewCol.Add cView, "VIEW" & cView.SQLViewName & cView.SQLViewOwner
    Next
    
    DB1COL.Add cPropOrg, cPropOrg.SQLName
    
    Set cPropOrg = New clsProps
    cPropOrg.SQLName = Me.cmbDatabases2.Text
    cPropOrg.SQLOwner = SqlSrv.Databases(Me.cmbDatabases2.Text).Owner
    
    Me.StatusBar1.SimpleText = "Retrieving stored procedures... " & Me.cmbDatabases2.Text
    For Each SQLObj In SqlSrv.Databases(Me.cmbDatabases2.Text).StoredProcedures
        If Not SQLObj.SystemObject Then
            Set cSP = New clsStorProcs
            cSP.SQLSPName = SQLObj.Name
            cSP.SQLSPCommand = SQLObj.Text
            cSP.SQLSPShortCmd = RemoveUnwantedChars(cSP.SQLSPCommand)
            cSP.SQLSPProcessed = False
            cSP.SQLSPOwner = SQLObj.Owner
            cSP.SQLSPID = SQLObj.ID
            cPropOrg.StorProcCol.Add cSP, "SP" & cSP.SQLSPName & cSP.SQLSPOwner
        End If
    Next
    
    Me.StatusBar1.SimpleText = "Retrieving tables... " & Me.cmbDatabases2.Text
    For Each SQLObj In SqlSrv.Databases(Me.cmbDatabases2.Text).Tables
        If Not SQLObj.SystemObject Then
            Set cTable = New clsTables
            cTable.SQLTableFieldsCount = SQLObj.Columns.Count
            cTable.SQLTableName = SQLObj.Name
            cTable.SQLTableProcessed = False
            cTable.SQLTableOwner = SQLObj.Owner
            cTable.SQLTableID = SQLObj.ID
            For Each SQLColumn In SQLObj.Columns
                Set cColumn = New clsColumns
                cColumn.SQLColIdentifier = SQLColumn.Identity
                cColumn.SQLColAllowNULL = SQLColumn.AllowNulls
                cColumn.SQLColName = SQLColumn.Name
                cColumn.SQLColType = SQLColumn.DataType
                cColumn.SQLColID = SQLColumn.ID
                cTable.ColumnCol.Add cColumn, cColumn.SQLColName
            Next
            cPropOrg.TableCol.Add cTable, "TABLE" & cTable.SQLTableName & cTable.SQLTableOwner
        End If
    Next
    
    Me.StatusBar1.SimpleText = "Retrieving views... " & Me.cmbDatabases2.Text
    For Each SQLObj In SqlSrv.Databases(Me.cmbDatabases2.Text).Views
        Set cView = New clsViews
        cView.SQLViewName = SQLObj.Name
        cView.SQLViewCommand = SQLObj.Text
        cView.SQLViewShortCmd = RemoveUnwantedChars(cView.SQLViewCommand)
        cView.SQLViewProcessed = False
        cView.SQLViewOwner = SQLObj.Owner
        cView.SQLViewID = SQLObj.ID
        cPropOrg.ViewCol.Add cView, "VIEW" & cView.SQLViewName & cView.SQLViewOwner
    Next
    
    DB1COL.Add cPropOrg, cPropOrg.SQLName
    
    Dim Node As MSComctlLib.Node
    Me.tvwSQL.Nodes.Clear
    
    Dim CompareRun As Boolean
    CompareRun = False
    
    Dim TekstOnObject As String
    
    For Each cPropOrg In DB1COL
        Set Node = Me.tvwSQL.Nodes.Add(, , cPropOrg.SQLName, cPropOrg.SQLName & " (" & cPropOrg.SQLOwner & ")")
        Node.ForeColor = vbRed
        Node.Expanded = True
        
        Set Node = Me.tvwSQL.Nodes.Add(cPropOrg.SQLName, tvwChild, "TABLE" & cPropOrg.SQLName, "Tables (" & cPropOrg.TableCol.Count & ")")
        Node.ForeColor = vbBlue
        Set Node = Me.tvwSQL.Nodes.Add(cPropOrg.SQLName, tvwChild, "SP" & cPropOrg.SQLName, "Stored Procedures (" & cPropOrg.StorProcCol.Count & ")")
        Node.ForeColor = vbBlue
        Set Node = Me.tvwSQL.Nodes.Add(cPropOrg.SQLName, tvwChild, "VIEW" & cPropOrg.SQLName, "Views (" & cPropOrg.ViewCol.Count & ")")
        Node.ForeColor = vbBlue
        
        For Each cSP In cPropOrg.StorProcCol
            TekstOnObject = ""
            Set Node = Me.tvwSQL.Nodes.Add("SP" & cPropOrg.SQLName, tvwChild)
            Node.Tag = cSP.SQLSPCommand
            If Not CompareRun Then
                If CompareOnObject("SP" & cSP.SQLSPName & cSP.SQLSPOwner, TekstOnObject, True) Then
                    Node.ForeColor = vbBlack
                Else
                    Node.ForeColor = vbRed
                End If
                Node.Text = "(" & cSP.SQLSPOwner & ") " & cSP.SQLSPName & " (" & TekstOnObject & ")"
            Else
                Node.Text = "(" & cSP.SQLSPOwner & ") " & cSP.SQLSPName
            End If
            
        Next
        
        For Each cTable In cPropOrg.TableCol
            TekstOnObject = ""
            Set Node = Me.tvwSQL.Nodes.Add("TABLE" & cPropOrg.SQLName, tvwChild)
            Node.Tag = cTable.SQLTableFieldsCount
            Node.Key = "COLUMN" & cPropOrg.SQLName & cTable.SQLTableName & cTable.SQLTableID
            If Not CompareRun Then
                If CompareOnObject("TABLE" & cTable.SQLTableName & cTable.SQLTableOwner, TekstOnObject, , True) Then
                    Node.ForeColor = vbBlack
                Else
                    Node.ForeColor = vbRed
                End If
                Node.Text = "(" & cTable.SQLTableOwner & ") " & cTable.SQLTableName & " (" & TekstOnObject & ")"
            Else
                Node.Text = "(" & cTable.SQLTableOwner & ") " & cTable.SQLTableName
            End If
            For Each cColumn In cTable.ColumnCol
                Set Node = Me.tvwSQL.Nodes.Add("COLUMN" & cPropOrg.SQLName & cTable.SQLTableName & cTable.SQLTableID, tvwChild)
                Node.Tag = cColumn.SQLColType
                Node.Text = cColumn.SQLColName & IIf(cColumn.SQLColIdentifier, " (ID)", "")
            Next
        Next
        
        For Each cView In cPropOrg.ViewCol
            TekstOnObject = ""
            Set Node = Me.tvwSQL.Nodes.Add("VIEW" & cPropOrg.SQLName, tvwChild)
            Node.Tag = cView.SQLViewCommand
            If Not CompareRun Then
                If CompareOnObject("VIEW" & cView.SQLViewName & cView.SQLViewOwner, TekstOnObject, , , True) Then
                    Node.ForeColor = vbBlack
                Else
                    Node.ForeColor = vbRed
                End If
                Node.Text = "(" & cView.SQLViewOwner & ") " & cView.SQLViewName & " (" & TekstOnObject & ")"
            Else
                Node.Text = "(" & cView.SQLViewOwner & ") " & cView.SQLViewName
            End If
        Next
        CompareRun = True
    Next
    
    
    Dim UpdOSQL As String
    
    If Me.optSQLVerif.Value Then
        UpdOSQL = "OSQL -S" & Me.txtServer.Text & " -U" & Me.txtLogin.Text & " -P " & Me.txtPassword.Text & " -i " & """" & App.Path & "\UpdateQuerys.sql" & """"
    Else
        UpdOSQL = "OSQL -E " & " -i " & """" & App.Path & "\UpdateQuerys.sql" & """"
    End If
    
    Me.txtScripting.Text = UpdOSQL
        
    If ScriptSettings.ScriptAutoProcess Then
        
        Dim Succeeded As Boolean
        
        ExecCmd UpdOSQL, Succeeded

    End If
    
    Screen.MousePointer = vbDefault
    
    Me.rtbDetails.SelStart = 0
    
    If ScriptingIsProcessed Then
        If ScriptSettings.ScriptAutoProcess Then
            Me.StatusBar1.SimpleText = "Major scripting done and processed"
        Else
            Me.StatusBar1.SimpleText = "Major scripting done.. copy/paste text in command window"
        End If
    Else
        Me.StatusBar1.SimpleText = "No major scripting done.."
    End If
    
    On Error Resume Next
    
    Me.txtScripting.SetFocus
    Me.txtScripting.SelStart = 0
    Me.txtScripting.SelLength = Len(Me.txtScripting.Text)
    
    On Error GoTo 0
    Exit Function
Err_Handler:
    Screen.MousePointer = vbDefault
    Me.StatusBar1.SimpleText = ""
    MsgBox Err.Description
    On Error GoTo 0
    Err.Clear
End Function

Private Function KillQueryFile()
    On Error Resume Next
    Kill App.Path & "\UpdateQuerys.sql"
End Function

Private Function CompareOnObject(ByVal ObjName As String, ByRef ObjT As String, Optional ByVal PS As Boolean = False, _
Optional ByVal PT As Boolean = False, Optional ByVal PV As Boolean = False) As Boolean
    On Error Resume Next
    If PS Then
        Dim CurS As clsStorProcs
        Dim ComS As clsStorProcs
        Set CurS = DB1COL.Item(1).StorProcCol(ObjName)
        Set ComS = DB1COL.Item(2).StorProcCol(ObjName)
        If (Not ComS Is Nothing) And (Not CurS Is Nothing) Then
            CurS.SQLSPProcessed = True
            ComS.SQLSPProcessed = True
            If StrComp(CurS.SQLSPShortCmd, ComS.SQLSPShortCmd, vbTextCompare) <> 0 Then
                ObjT = "COMMAND"
                WriteLogOnRTB "COMMAND > Stored Procedure " & CurS.SQLSPName & " ON " & DB1COL.Item(1).SQLName, CurS.SQLSPCommand
                GenerateScripting CurS.SQLSPName, enmScriptStoredProcedure
                CompareOnObject = False
            Else
                ObjT = "OK"
                CompareOnObject = True
            End If
        Else
            If Not CurS Is Nothing Then
                If Not ScriptSettings.ScriptStoredProcedures Then Exit Function
                If Not ScriptSettings.ScriptOwnerDiff Then
                    If TraceObjectInCol(CurS.SQLSPName, enmScriptStoredProcedure) Then
                        ObjT = "Not scripted"
                        CompareOnObject = True
                        Exit Function
                    End If
                End If
                CurS.SQLSPProcessed = True
                ObjT = "MISSING"
                WriteLogOnRTB "MISSING > Stored Procedure " & CurS.SQLSPName & " ON " & DB1COL.Item(1).SQLName, CurS.SQLSPCommand
                GenerateScripting CurS.SQLSPOwner & "." & CurS.SQLSPName, enmScriptStoredProcedure
            Else
                If Not ComS Is Nothing Then
                    ComS.SQLSPProcessed = True
                    ObjT = "MISSING"
                    WriteLogOnRTB "MISSING > Stored Procedure " & ComS.SQLSPName & " ON " & DB1COL.Item(2).SQLName, ComS.SQLSPCommand
                Else
                    ' cant happen.. but who cares...
                    WriteLogOnRTB "MISSING > Stored Procedure on both databases", ""
                End If
            End If
            CompareOnObject = False
        End If
    ElseIf PT Then
        Dim CurT As clsTables
        Dim ComT As clsTables
        Dim ComC As clsColumns
        Dim CurC As clsColumns
        Set CurT = DB1COL.Item(1).TableCol(ObjName)
        Set ComT = DB1COL.Item(2).TableCol(ObjName)
        If (Not ComT Is Nothing) And (Not CurT Is Nothing) Then
            CurT.SQLTableProcessed = True
            ComT.SQLTableProcessed = True
            If CurT.SQLTableFieldsCount <> ComT.SQLTableFieldsCount Then
                If ScriptSettings.ScriptColumns Then
                    Dim ColMsg As String
                    ObjT = "COUNT"
                    CompareOnObject = False
                    On Error Resume Next
                    For Each CurC In CurT.ColumnCol
                        Set ComC = ComT.ColumnCol(CurC.SQLColName & CurC.SQLColID)
                        If StrComp(ComC.SQLColName, CurC.SQLColName, vbTextCompare) <> 0 Then
                            ColMsg = ColMsg & vbCrLf & CurC.SQLColName & " (" & CurC.SQLColType & ")"
                            GenerateScripting "", enmScriptAlterTable, "ALTER TABLE " & CurT.SQLTableName & " ADD " & CurC.SQLColName & " " & CurC.SQLColType & IIf(CurC.SQLColAllowNULL, " NULL", "")
                        End If
                    Next
                    WriteLogOnRTB "COUNT > Fields ON Table " & CurT.SQLTableName, "COUNT = " & CurT.SQLTableFieldsCount & " ON " & DB1COL.Item(1).SQLName & ColMsg
                End If
            Else
                ObjT = "OK"
                CompareOnObject = True
            End If
        Else
            If Not CurT Is Nothing Then
                If Not ScriptSettings.ScriptTables Then Exit Function
                If Not ScriptSettings.ScriptOwnerDiff Then
                    If TraceObjectInCol(CurT.SQLTableName, enmScriptTable) Then
                        ObjT = "Not scripted"
                        CompareOnObject = True
                        Exit Function
                    End If
                End If
                CurT.SQLTableProcessed = True
                ObjT = "MISSING"
                GenerateScripting CurT.SQLTableOwner & "." & CurT.SQLTableName, enmScriptTable
                WriteLogOnRTB "MISSING > Table or FieldCount " & CurT.SQLTableName, "Table count " & CurT.SQLTableFieldsCount & " ON " & DB1COL.Item(1).SQLName
            Else
                If Not ComT Is Nothing Then
                    ComT.SQLTableProcessed = True
                    ObjT = "MISSING"
                    WriteLogOnRTB "MISSING > Table or FieldCount " & ComT.SQLTableName, "Table count " & ComT.SQLTableFieldsCount & " ON " & DB1COL.Item(2).SQLName
                Else
                    ' cant happen.. but who cares...
                    WriteLogOnRTB "MISSING > Table or FieldCount on both databases", ""
                End If
            End If
            CompareOnObject = False
        End If
    ElseIf PV Then
        Dim CurV As clsViews
        Dim ComV As clsViews
        Set CurV = DB1COL.Item(1).ViewCol(ObjName)
        Set ComV = DB1COL.Item(2).ViewCol(ObjName)
        If (Not ComV Is Nothing) And (Not CurV Is Nothing) Then
            CurV.SQLViewProcessed = True
            ComV.SQLViewProcessed = True
            If StrComp(CurV.SQLViewShortCmd, ComV.SQLViewShortCmd, vbTextCompare) <> 0 Then
                ObjT = "COMMAND"
                WriteLogOnRTB "COMMAND > View " & CurV.SQLViewName & " ON " & DB1COL.Item(1).SQLName, CurV.SQLViewCommand
                GenerateScripting CurV.SQLViewName, enmScriptView
                CompareOnObject = False
            Else
                ObjT = "OK"
                CompareOnObject = True
            End If
        Else
            If Not CurV Is Nothing Then
                If Not ScriptSettings.ScriptViews Then Exit Function
                If Not ScriptSettings.ScriptOwnerDiff Then
                    If TraceObjectInCol(CurV.SQLViewName, enmScriptView) Then
                        ObjT = "Not scripted"
                        CompareOnObject = True
                        Exit Function
                    End If
                End If
                CurV.SQLViewProcessed = True
                ObjT = "MISSING"
                GenerateScripting CurV.SQLViewOwner & "." & CurV.SQLViewName, enmScriptView
                WriteLogOnRTB "MISSING > View " & CurV.SQLViewName & " ON " & DB1COL.Item(1).SQLName, CurV.SQLViewCommand
            Else
                If Not ComV Is Nothing Then
                    ComV.SQLViewProcessed = True
                    ObjT = "MISSING"
                    WriteLogOnRTB "MISSING > View " & ComV.SQLViewName & " ON " & DB1COL.Item(2).SQLName, ComV.SQLViewCommand
                Else
                ' still cant happen.. but .. you know it
                    WriteLogOnRTB "MISSING > View on both databases", ""
                End If
            End If
            CompareOnObject = False
        End If
    End If
End Function

Private Function TraceObjectInCol(ByVal ObjName As String, ByVal SO As ScriptObjects) As Boolean
    If DB1COL Is Nothing Then Exit Function
    If DB1COL.Count < 2 Then Exit Function
    TraceObjectInCol = False
    Select Case SO
        Case enmScriptStoredProcedure
            Dim ObjSP As clsStorProcs
            For Each ObjSP In DB1COL.Item(2).StorProcCol
                If UCase(ObjSP.SQLSPName) = UCase(ObjName) Then
                    TraceObjectInCol = True
                    Exit Function
                End If
            Next
        Case enmScriptAlterTable
        Case enmScriptTable
            Dim ObjTB As clsTables
            For Each ObjTB In DB1COL.Item(2).TableCol
                If UCase(ObjTB.SQLTableName) = UCase(ObjName) Then
                    TraceObjectInCol = True
                    Exit Function
                End If
            Next
        Case enmScriptView
            Dim ObjVW As clsViews
            For Each ObjVW In DB1COL.Item(2).ViewCol
                If UCase(ObjVW.SQLViewName) = UCase(ObjName) Then
                    TraceObjectInCol = True
                    Exit Function
                End If
            Next
    End Select
    
End Function

Public Function RemoveUnwantedChars(ByVal T As String) As String
    T = Replace(T, " ", "")
    T = Replace(T, vbCr, "")
    T = Replace(T, vbLf, "")
    T = Replace(T, vbCrLf, "")
    RemoveUnwantedChars = T
End Function

Private Function WriteLogOnRTB(ByVal StrHeader As String, ByVal StrText As String)
    With Me.rtbDetails
        .SelStart = Len(.Text)
        .SelColor = vbRed
        .SelUnderline = True
        .SelBold = True
        .SelText = vbCrLf & StrHeader
        .SelStart = Len(.Text)
        .SelColor = vbBlack
        .SelUnderline = False
        .SelBold = False
        .SelText = vbCrLf & StrText & vbCrLf & vbCrLf
    End With
End Function

Private Sub GenerateScripting(ByVal T As String, ByVal ScriptType As ScriptObjects, Optional ByVal AlterTable As String = "")
    On Error GoTo Err_Handler
    
    Dim ScriptObj As Object
    
    ScriptingIsProcessed = True
    
    If ScriptType = enmScriptStoredProcedure Then
    
        If Not ScriptSettings.ScriptStoredProcedures Then Exit Sub
    
        iScriptType = SQLDMOScript_Drops _
                      Or SQLDMOScript_ObjectPermissions _
                      Or SQLDMOScript_OwnerQualify _
                      Or SQLDMOScript_Default
        iScript2Type = SQLDMOScript2_Default
    
        Set ScriptObj = SqlSrv.Databases(Me.cmbDatabases1.Text).StoredProcedures(T)
    
    ElseIf ScriptType = enmScriptView Then
    
        If Not ScriptSettings.ScriptViews Then Exit Sub
    
        iScriptType = SQLDMOScript_Drops _
                      Or SQLDMOScript_ObjectPermissions _
                      Or SQLDMOScript_OwnerQualify _
                      Or SQLDMOScript_Default
        iScript2Type = SQLDMOScript2_Default
    
        Set ScriptObj = SqlSrv.Databases(Me.cmbDatabases1.Text).Views(T)
    
    ElseIf ScriptType = enmScriptTable Then
    
        If Not ScriptSettings.ScriptTables Then Exit Sub
    
        iScriptType = SQLDMOScript_ObjectPermissions _
                      Or SQLDMOScript_OwnerQualify _
                      Or SQLDMOScript_Default _
                      Or SQLDMOScript_Indexes _
                      Or SQLDMOScript_DRI_All
        iScript2Type = SQLDMOScript2_NoWhatIfIndexes
        
        Set ScriptObj = SqlSrv.Databases(Me.cmbDatabases1.Text).Tables(T)
        
    ElseIf ScriptType = enmScriptAlterTable Then
        
        If Not ScriptSettings.ScriptColumns Then Exit Sub
        
    End If
    
    Dim Fi As Long
    If ScriptType = enmScriptAlterTable Then

            Fi = FreeFile
            Open App.Path & "\UpdateQuerys.sql" For Append As Fi
            
            Print #Fi, "USE " & Me.cmbDatabases2.Text
            Print #Fi, ""
            Print #Fi, "GO"
            Print #Fi, ""
    
            Print #Fi, AlterTable
            Print #Fi, ""
            Close Fi
    Else
    
        If Not ScriptObj Is Nothing Then
        
            Fi = FreeFile
            Open App.Path & "\UpdateQuerys.sql" For Append As Fi
            
            Print #Fi, "USE " & Me.cmbDatabases2.Text
            Print #Fi, ""
            Print #Fi, "GO"
            Print #Fi, ""
            
            If ScriptType = enmScriptTable Then
            
                Print #Fi, ScriptObj.Script(iScriptType, , , iScript2Type)
            
            Else
                
                Print #Fi, ScriptObj.Script(iScriptType, , iScript2Type)
                
            End If
            Print #Fi, ""
            Close Fi
        
        End If
        
    End If
    
    If Not ScriptObj Is Nothing Then Set ScriptObj = Nothing
    
    On Error GoTo 0
    Exit Sub
Err_Handler:
    If Not ScriptObj Is Nothing Then Set ScriptObj = Nothing
    Reset
    Err.Clear
    On Error GoTo 0
End Sub


Private Sub ResizeControlsIR()
    Me.picSplitter.Top = Me.picResizer.Top
    Me.picResizer.Visible = False
    Me.tvwSQL.Move 0, 10, Me.picBG.Width - 20, Me.picSplitter.Top - 20
    Me.rtbDetails.Move 0, Me.tvwSQL.Height + 120, Me.picBG.Width - 20, Me.picBG.Height - Me.tvwSQL.Height - 140
End Sub

Private Sub PicSplitter_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    If Button = vbLeftButton Then
        bSplitStarted = True
        Me.picResizer.Top = Me.picSplitter.Top
        iSplitterX = Me.picResizer.Top
    End If
End Sub

Private Sub PicSplitter_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    Dim L As Single
    If bSplitStarted Then
        Me.picResizer.Visible = True
        Me.picResizer.ZOrder 0
        If (y + iSplitterX) > Me.picBG.Height - 500 Then Exit Sub
        If (y + iSplitterX) < 500 Then Exit Sub
        Me.picResizer.Top = y + iSplitterX
        iCurrentX = Me.picResizer.Top
    End If
End Sub

Private Sub picSplitter_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    bSplitStarted = False
    ResizeControlsIR
End Sub
