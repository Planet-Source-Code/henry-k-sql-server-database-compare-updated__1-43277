VERSION 5.00
Begin VB.Form frmSettings 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Settings"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4305
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   4305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraSettings 
      Height          =   3015
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   3015
      Begin VB.CheckBox chkAutorun 
         Caption         =   "Automatically run script"
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
         Left            =   120
         TabIndex        =   7
         Top             =   2640
         Width           =   2535
      End
      Begin VB.CheckBox chkProcedures 
         Caption         =   "Script missing stored procedures"
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
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   2775
      End
      Begin VB.CheckBox chkViews 
         Caption         =   "Script missing views"
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
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   2775
      End
      Begin VB.CheckBox chkTables 
         Caption         =   "Script missing tables"
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
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   2775
      End
      Begin VB.CheckBox chkColumns 
         Caption         =   "Script missing columns"
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
         Left            =   120
         TabIndex        =   3
         Top             =   1680
         Width           =   2535
      End
      Begin VB.CheckBox chkOwner 
         Caption         =   "Script if owner is different"
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
         Left            =   120
         TabIndex        =   2
         Top             =   2160
         Width           =   2535
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   3120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub FormDoModal()
    Me.chkProcedures.Value = IIf(modMain.ScriptSettings.ScriptStoredProcedures, vbChecked, vbUnchecked)
    Me.chkViews.Value = IIf(modMain.ScriptSettings.ScriptViews, vbChecked, vbUnchecked)
    Me.chkTables.Value = IIf(modMain.ScriptSettings.ScriptTables, vbChecked, vbUnchecked)
    Me.chkColumns.Value = IIf(modMain.ScriptSettings.ScriptColumns, vbChecked, vbUnchecked)
    Me.chkOwner.Value = IIf(modMain.ScriptSettings.ScriptOwnerDiff, vbChecked, vbUnchecked)
    Me.chkAutorun.Value = IIf(modMain.ScriptSettings.ScriptAutoProcess, vbChecked, vbUnchecked)
    Me.Show vbModal
End Sub

Private Sub chkAutorun_Click()
    ScriptSettings.ScriptAutoProcess = Me.chkAutorun.Value = vbChecked
End Sub

Private Sub chkColumns_Click()
    ScriptSettings.ScriptColumns = Me.chkColumns.Value = vbChecked
End Sub

Private Sub chkOwner_Click()
    ScriptSettings.ScriptOwnerDiff = Me.chkOwner.Value = vbChecked
End Sub

Private Sub chkProcedures_Click()
    ScriptSettings.ScriptStoredProcedures = Me.chkProcedures.Value = vbChecked
End Sub

Private Sub chkTables_Click()
    ScriptSettings.ScriptTables = Me.chkTables.Value = vbChecked
End Sub

Private Sub chkViews_Click()
    ScriptSettings.ScriptViews = Me.chkViews.Value = vbChecked
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub
