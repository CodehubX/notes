VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmZoeken 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Zoeken"
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6255
   HelpContextID   =   1140
   Icon            =   "frmZoeken.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   6255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   2535
      Left            =   120
      TabIndex        =   9
      Top             =   30
      Width           =   6015
      Begin VB.CheckBox chkSearchDescription 
         Caption         =   "Beschrijving"
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
         Left            =   960
         TabIndex        =   5
         Top             =   1680
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CheckBox chkSearchTrefwoorden 
         Caption         =   "Trefwoorden"
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
         Left            =   960
         TabIndex        =   4
         Top             =   1440
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.TextBox txtZoekWoorden 
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
         Left            =   1200
         TabIndex        =   0
         Top             =   240
         Width           =   4695
      End
      Begin VB.CheckBox chkHeelWoord 
         Caption         =   "Hele woorden"
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
         Top             =   720
         Width           =   1335
      End
      Begin VB.CheckBox chkIdentiekeLetters 
         Caption         =   "Identieke hoofdletters/kleine letters"
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
         Top             =   960
         Width           =   2895
      End
      Begin VB.CommandButton cmdZoek 
         Caption         =   "&Zoeken"
         Default         =   -1  'True
         Enabled         =   0   'False
         Height          =   375
         Left            =   4800
         TabIndex        =   1
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Zoeken in"
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
         TabIndex        =   13
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Zoeken naar"
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
         TabIndex        =   11
         Top             =   270
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Gebruik een komma om de woorden te scheiden als u op meer dan één woord wilt zoeken."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   120
         TabIndex        =   10
         Top             =   2040
         Width           =   5775
      End
   End
   Begin VB.Frame Frame2 
      Height          =   3375
      Left            =   120
      TabIndex        =   12
      Top             =   2550
      Width           =   6015
      Begin MSComctlLib.ListView lvwResults 
         Height          =   3015
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   5318
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "Column1"
            Text            =   "Notes"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Key             =   "Column2"
            Text            =   "Gevonden Woorden"
            Object.Width           =   4410
         EndProperty
      End
   End
   Begin VB.CommandButton cmdGoTo 
      Caption         =   "&Ga naar note"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3720
      TabIndex        =   7
      Top             =   6030
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   5040
      TabIndex        =   8
      Top             =   6030
      Width           =   1095
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      X1              =   -15
      X2              =   10000
      Y1              =   15
      Y2              =   15
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   -15
      X2              =   10000
      Y1              =   0
      Y2              =   0
   End
End
Attribute VB_Name = "frmZoeken"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------
'Copyright © 1998 - 2004 Wout Maaskant
'
'This file is part of WMS Notes.
'
'WMS Notes is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'(at your option) any later version.
'
'WMS Notes is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with WMS Notes; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'--------------------------------------------------------------------

Private Sub cmdGoTo_Click()
    On Error GoTo cmdGoTOError
    Dim NodX As Node

    'Select: Node
    Set NodX = frmMain.lstItems.Nodes.Item(lvwResults.SelectedItem.Key)
    NodX.EnsureVisible
    NodX.Selected = True
    NodeChange NodX

    'Update: frmZoeken
    frmZoeken.Hide

    Exit Sub
cmdGoTOError:
    Err.Clear
End Sub
Private Sub cmdOK_Click()
    Unload Me
End Sub
Private Sub cmdZoek_Click()
    lvwResults.ListItems.Clear
    cmdGoTo.Enabled = False

    'Zoeken.
    Zoeken txtZoekWoorden.Text, chkSearchTrefwoorden.Value, chkSearchDescription.Value, chkIdentiekeLetters.Value, chkHeelWoord.Value
End Sub
Private Sub Form_Load()
    lvwResults.ColumnHeaders.Item("Column1").Width = 2300
    lvwResults.ColumnHeaders.Item("Column2").Width = 3200
End Sub
Private Sub lvwResults_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lvwResults.SortKey = ColumnHeader.Index - 1
End Sub
Private Sub lvwResults_DblClick()
    If cmdGoTo.Enabled Then cmdGoTo_Click
End Sub
Private Sub lvwResults_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If Item.Key = "Nothing" Then
        cmdGoTo.Enabled = False
    Else
        cmdGoTo.Enabled = True
    End If
End Sub
Private Sub txtZoekWoorden_Change()
    Dim ZoekTekst As String

    'Controleer de zoektekst.
    ZoekTekst = txtZoekWoorden.Text
    If ZoekTekst = "" Then
        cmdZoek.Enabled = False
    ElseIf Right(ZoekTekst, 1) = "," Then
        cmdZoek.Enabled = False
    ElseIf InStr(1, ZoekTekst, ",,") <> 0 Then
        cmdZoek.Enabled = False
    Else
        cmdZoek.Enabled = True
    End If
End Sub
