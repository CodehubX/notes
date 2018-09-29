VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmNoteInfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Note Eigenschappen"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4935
   HelpContextID   =   1140
   Icon            =   "frmNoteInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton cmdAnnuleren 
      Caption         =   "Annuleren"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Frame framePropInfo 
      BorderStyle     =   0  'None
      Height          =   3375
      Left            =   240
      TabIndex        =   3
      Top             =   480
      Width           =   4455
      Begin VB.TextBox txtNoteNummer 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
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
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "NOTENUMMER"
         Top             =   840
         Width           =   3255
      End
      Begin VB.TextBox txtNoteTrefwoorden 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1755
         Left            =   1200
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   9
         Text            =   "frmNoteInfo.frx":0E42
         Top             =   1560
         Width           =   3255
      End
      Begin VB.TextBox txtNoteGroep 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
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
         Locked          =   -1  'True
         TabIndex        =   8
         Text            =   "NOTEGROEP"
         Top             =   1200
         Width           =   3255
      End
      Begin VB.TextBox txtNoteNaam 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   720
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   5
         Text            =   "frmNoteInfo.frx":0E52
         Top             =   240
         Width           =   3375
      End
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   120
         Picture         =   "frmNoteInfo.frx":0E5C
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   4
         Top             =   120
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Notenummer: "
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
         TabIndex        =   10
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Trefwoorden: "
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
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Groep: "
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
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000014&
         X1              =   120
         X2              =   4320
         Y1              =   735
         Y2              =   735
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         X1              =   120
         X2              =   4320
         Y1              =   720
         Y2              =   720
      End
   End
   Begin VB.Frame framePropZoek 
      BorderStyle     =   0  'None
      Height          =   3375
      Left            =   240
      TabIndex        =   12
      Top             =   480
      Visible         =   0   'False
      Width           =   4455
      Begin VB.TextBox txtZoekWoord 
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
         Left            =   2520
         TabIndex        =   16
         Top             =   360
         Width           =   1935
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "Trefwoord &verwijderen"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2520
         TabIndex        =   15
         Top             =   1200
         Width           =   1935
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Trefwoord &toevoegen"
         Default         =   -1  'True
         Enabled         =   0   'False
         Height          =   375
         Left            =   2520
         TabIndex        =   14
         Top             =   720
         Width           =   1935
      End
      Begin VB.ListBox lstTermen 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2940
         IntegralHeight  =   0   'False
         Left            =   0
         Sorted          =   -1  'True
         TabIndex        =   13
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label4 
         Caption         =   "Trefwoorden:"
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
         Left            =   0
         TabIndex        =   17
         Top             =   120
         Width           =   1215
      End
   End
   Begin MSComctlLib.TabStrip tabTabStrip 
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   6800
      TabMinWidth     =   0
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Note informatie"
            Key             =   "PropInfo"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Trefwoorden wijzigen"
            Key             =   "PropZoek"
            ImageVarType    =   2
         EndProperty
      EndProperty
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
   Begin VB.Line Line4 
      BorderColor     =   &H80000010&
      X1              =   -15
      X2              =   10000
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000014&
      X1              =   -15
      X2              =   10000
      Y1              =   15
      Y2              =   15
   End
End
Attribute VB_Name = "frmNoteInfo"
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

Public Canceled As Boolean
Private Sub cmdAdd_Click()
    Dim ZoekTerm As String

    'Update: Controls
    lstTermen.AddItem Trim(txtZoekWoord.Text)
    txtZoekWoord.Text = ""
    cmdAdd.Enabled = False
    cmdAnnuleren.Enabled = True
End Sub
Private Sub cmdAnnuleren_Click()
    Me.Canceled = True
    Me.Hide
End Sub
Private Sub cmdOK_Click()
    If cmdAdd.Enabled = True Then cmdAdd_Click
    Me.Canceled = False
    Me.Hide
End Sub
Private Sub cmdRemove_Click()
    'Update: Controls
    lstTermen.RemoveItem lstTermen.ListIndex
    cmdRemove.Enabled = False
    cmdAnnuleren.Enabled = True
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'Als er via het menu/kruisje afgesloten wordt -> Annuleren.
    If UnloadMode = vbFormControlMenu Then
        Cancel = 1
        cmdAnnuleren_Click
    End If
End Sub
Private Sub lstTermen_Click()
    If lstTermen.ListIndex = -1 Then cmdRemove.Enabled = False Else cmdRemove.Enabled = True
End Sub
Private Sub txtZoekWoord_Change()
    Dim ZoekWoord As String
    Dim cntTermen As Integer, iTerm As Integer

    'Init: ZoekWoord
    ZoekWoord = txtZoekWoord.Text

    'Check: ZoekWoord
    If Trim(ZoekWoord) = "" Then
        cmdAdd.Enabled = False
        Exit Sub
    End If

    'Check: ZoekWoord
    'Kijk of het trefwoord al bestaat.
    cntTermen = lstTermen.ListCount
    For iTerm = 0 To cntTermen - 1
        If lstTermen.List(iTerm) = ZoekWoord Then
            cmdAdd.Enabled = False
            Exit Sub
        End If
    Next

    'Update: cmdAdd
    If Not cmdAdd.Enabled Then cmdAdd.Enabled = True
End Sub
Private Sub txtZoekWoord_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) = "," Then KeyAscii = 0
End Sub
Private Sub tabTabStrip_Click()
    If tabTabStrip.SelectedItem.Key = "PropInfo" Then
        framePropInfo.Visible = True
        framePropZoek.Visible = False
    ElseIf tabTabStrip.SelectedItem.Key = "PropZoek" Then
        framePropInfo.Visible = False
        framePropZoek.Visible = True
    End If
End Sub
