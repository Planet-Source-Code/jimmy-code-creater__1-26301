VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Create 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "DB Coder"
   ClientHeight    =   7245
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5790
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7245
   ScaleWidth      =   5790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox formnametextbox 
      Height          =   285
      Left            =   150
      TabIndex        =   7
      Top             =   5730
      Width           =   3045
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Make like a &Tree"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   120
      TabIndex        =   6
      Top             =   6630
      Width           =   5370
   End
   Begin VB.ListBox tablelist 
      Height          =   3570
      Left            =   150
      TabIndex        =   5
      ToolTipText     =   "Select the tabel to use"
      Top             =   1755
      Width           =   2730
   End
   Begin VB.ListBox fieldlist 
      Height          =   3570
      Left            =   3060
      MultiSelect     =   1  'Simple
      TabIndex        =   4
      ToolTipText     =   "Select the fields you want the code created for"
      Top             =   1785
      Width           =   2385
   End
   Begin VB.TextBox databasepathtextbox 
      Height          =   285
      Left            =   150
      TabIndex        =   3
      Top             =   510
      Width           =   4935
   End
   Begin VB.CommandButton dbpathbutton 
      Caption         =   "..."
      Height          =   270
      Left            =   5085
      TabIndex        =   2
      Top             =   525
      Width           =   540
   End
   Begin VB.CommandButton loadbutton 
      Caption         =   "Load Tables"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   180
      TabIndex        =   1
      Top             =   975
      Width           =   2760
   End
   Begin VB.CommandButton createcode 
      Caption         =   "&Code it"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   105
      TabIndex        =   0
      Top             =   6120
      Width           =   5385
   End
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   3450
      Top             =   915
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Database path"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   210
      TabIndex        =   11
      Top             =   270
      Width           =   1155
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Field List"
      Height          =   195
      Left            =   3120
      TabIndex        =   10
      Top             =   1500
      Width           =   615
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Table List"
      Height          =   195
      Left            =   195
      TabIndex        =   9
      Top             =   1485
      Width           =   690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Form name"
      Height          =   195
      Left            =   165
      TabIndex        =   8
      Top             =   5430
      Width           =   780
   End
End
Attribute VB_Name = "Create"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim conDB As ADODB.Connection
Dim rstTables As ADODB.Recordset

Private Sub CancelButton_Click()
    End
End Sub

Private Sub createcode_Click()

On Error GoTo Log

Dim listcount As Integer

Dim filehandle As Long
Dim location As String
Dim tops As Long

If formnametextbox.Text = "" Then
    MsgBox "What is this form going to be called?"
    formnametextbox.SetFocus
    Exit Sub
End If

If fieldlist.listcount = 0 Then
    MsgBox "There or no fields to use"
    Exit Sub
End If

filehandle = FreeFile

location = GetFilePath(databasepathtextbox) & formnametextbox & ".frm"

Open location For Output As filehandle

    Print #filehandle, "Version 5.00"
    Print #filehandle, "Begin VB.Form " & formnametextbox.Text
    Print #filehandle, "AutoRedraw = -1           'True"
    Print #filehandle, "BackColor = &HC0C0C0"
    Print #filehandle, "BorderStyle = 3        'Fixed Dialog"
    Print #filehandle, "ClientHeight = 3135"
    Print #filehandle, "ClientLeft = 45"
    Print #filehandle, "ClientTop = 330"
    Print #filehandle, "ClientWidth = 7650"
    'Print #filehandle, "Icon            =   " & formnametextbox.Text & ".frx" & ":0000"
    Print #filehandle, "LinkTopic =   " & """" & "Form1" & """"
    Print #filehandle, "MaxButton = 0           'False"
    Print #filehandle, "MDIChild = 0           'True"
    Print #filehandle, "MinButton = 0           'False"
    Print #filehandle, "ScaleHeight = 5000"
    Print #filehandle, "ScaleWidth = 7650"
    Print #filehandle, "ShowInTaskbar = 0       'False"

      tops = 100
For listcount = 0 To fieldlist.listcount - 1
    fieldlist.ListIndex = listcount
    If fieldlist.Selected(listcount) = True Then
        tops = tops + listcount + 385
        Print #filehandle, "Begin VB.TextBox " & fieldlist.Text & "textbox"
        Print #filehandle, "    Height          =   300"
        Print #filehandle, "    Left             =   5355"
        Print #filehandle, "    TabIndex      =   " & listcount
        Print #filehandle, "    Top             =   " & tops
        Print #filehandle, "    Width           =   1965"
        Print #filehandle, "End"
        Print #filehandle, "Begin VB.Label " & fieldlist.Text & "label"
        Print #filehandle, "    AutoSize = -1          'True"
        Print #filehandle, "    BackStyle = 0           'Transparent"
        Print #filehandle, "    Caption = " & """" & fieldlist.Text & """"
        Print #filehandle, "    Height = 330"
        Print #filehandle, "    Left = 615"
        Print #filehandle, "    TabIndex = 1"
        Print #filehandle, "    Top = " & tops; ""
        Print #filehandle, "    Width = 2820"
        Print #filehandle, "End"
    End If
Next listcount

Print #filehandle, "Begin VB.CommandButton savebutton"
Print #filehandle, "Caption = " & """" & "&Save" & """"
Print #filehandle, "BeginProperty Font"
Print #filehandle, "Name = " & """" & "Arial" & """"
Print #filehandle, "Size = 9"
Print #filehandle, "Charset = 0"
Print #filehandle, "Weight = 700"
Print #filehandle, "Underline = 0                   'False"
Print #filehandle, "Italic = 0                      'False"
Print #filehandle, "Strikethrough = 0               'False"
Print #filehandle, "EndProperty"
Print #filehandle, "Height = 360"
Print #filehandle, "Left = 750"
Print #filehandle, "TabIndex = 3"
Print #filehandle, "Top = 2805"
Print #filehandle, "Width = 3000"
Print #filehandle, "End"
Print #filehandle, "Begin VB.CommandButton cancelbutton"
Print #filehandle, "Caption = " & """" & "&Cancel" & """"
Print #filehandle, " BeginProperty Font"
Print #filehandle, "Name = " & """" & "Arial" & """"
Print #filehandle, "Size = 9"
Print #filehandle, "Charset = 0"
Print #filehandle, "Weight = 700"
Print #filehandle, "Underline = 0                   'False"
Print #filehandle, "Italic = 0                      'False"
Print #filehandle, "Strikethrough = 0               'False"
Print #filehandle, "EndProperty"
Print #filehandle, "Height = 360"
Print #filehandle, "Left = 645"
Print #filehandle, "TabIndex = 2"
Print #filehandle, "Top = 2205"
Print #filehandle, "Width = 2805"
Print #filehandle, "End"
Print #filehandle, "End"
Print #filehandle, "Attribute VB_Name = " & """" & formnametextbox.Text & """"
Print #filehandle, "Attribute VB_GlobalNameSpace = False"
Print #filehandle, "Attribute VB_Creatable = False"
Print #filehandle, "Attribute VB_PredeclaredId = True"
Print #filehandle, "Attribute VB_Exposed = False"

Print #filehandle, "Option Explicit"

Print #filehandle, "Public sub Savebutton_ClicK"
Print #filehandle, "Dim tb" & tablelist.Text & " As ADODB.Recordset"
Print #filehandle, "Set tb" & tablelist.Text & " = new Recordset"
Print #filehandle, "tb" & tablelist.Text & ".Addnew"
For listcount = 0 To fieldlist.listcount - 1
    If fieldlist.Selected(listcount) = True Then
        fieldlist.ListIndex = listcount
        tops = tops + listcount + 385
        Print #filehandle, "    tb" & tablelist.Text & ".Fields(" & """" & fieldlist.Text & """" & ") = " & "Nullit(" & fieldlist.Text & "textbox" & ")"
    End If
Next listcount
Print #filehandle, "tb" & tablelist.Text & ".Update"
Print #filehandle, "tb" & tablelist.Text & ".close"
Print #filehandle, "set tb" & tablelist.Text & ".Activeconnection = Nothing"
Print #filehandle, "End Sub"
    

Print #filehandle, "Private Sub cancelbutton_click"
Print #filehandle, " Unload me"
Print #filehandle, "End Sub"

Print #filehandle, "Function Nullit(ctl as Control) As Variant"
Print #filehandle, "Select case ctl"
Print #filehandle, "   Case """""""
Print #filehandle, "       Nullit = Null"
Print #filehandle, "    Case Else"
Print #filehandle, "       Nullit = ctl"
Print #filehandle, "End Select"
Print #filehandle, "End Function"


Close #filehandle

MsgBox "Code is now done", vbInformation

Exit Sub

Log:
MsgBox Err.Description & " " & Err.Number, vbCritical

End Sub

Private Sub dbpathbutton_Click()
On Error GoTo Log

Dialog.Filter = "Database files (*.mdb)|*.mdb| "
Dialog.ShowOpen
databasepathtextbox.Text = Dialog.FileName

Exit Sub

Log:
MsgBox Err.Description & " " & Err.Number, vbCritical
End Sub

Private Sub loadbutton_Click()

On Error GoTo Log
    
    Set conDB = New ADODB.Connection
    With conDB
                .Provider = "Microsoft.Jet.OLEDB.4.0;" 'Data Source=%DB%;Persist Security Info=False"
                .ConnectionString = databasepathtextbox.Text
                .Open
    End With
    
    Set rstTables = conDB.OpenSchema(adSchemaTables)
    
    While Not rstTables.EOF
        tablelist.AddItem rstTables.Fields(2)
        rstTables.MoveNext
    Wend

Exit Sub

Log:
MsgBox Err.Description & " " & Err.Number, vbCritical
End Sub


Private Sub tablelist_Click()

On Error GoTo Log

Dim tables As ADODB.Recordset
Dim fieldname As ADODB.Field

fieldlist.Clear
Set tables = New Recordset
tables.Open (tablelist.Text), conDB, adOpenForwardOnly, adLockReadOnly

    For Each fieldname In tables.Fields
        fieldlist.AddItem fieldname.Name
    Next
formnametextbox.Text = UCase$(tablelist.Text) & "_Form"
Exit Sub

Log:
MsgBox Err.Description & " " & Err.Number, vbCritical

End Sub

Public Function GetFilePath(FileNamePath As String) As String
    On Error GoTo FunctionError:
    Dim x
    Dim tString As String
    'gets the file's path from the file


    For x = Len(FileNamePath) To 0 Step -1
        tString = Mid$(FileNamePath, x, 1)


        If tString = "\" Then
            GetFilePath = Left(FileNamePath, x)
            Exit Function
        End If
    Next x


FunctionError:
    GetFilePath = -1
End Function
