VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSetup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Opties"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7815
   Icon            =   "frmSetup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   7815
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picHyperlinks 
      BackColor       =   &H00FFFFFF&
      Height          =   3015
      Left            =   120
      ScaleHeight     =   2955
      ScaleWidth      =   2115
      TabIndex        =   2
      Top             =   120
      Width           =   2175
      Begin VB.PictureBox picHyperlink 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   2
         Left            =   345
         ScaleHeight     =   255
         ScaleWidth      =   1575
         TabIndex        =   36
         Top             =   840
         Width           =   1575
         Begin VB.Timer tmrHyperlink 
            Index           =   2
            Left            =   0
            Top             =   120
         End
         Begin VB.Label lblHyperlink 
            AutoSize        =   -1  'True
            Caption         =   "Bestandskoppeling"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   0
            MouseIcon       =   "frmSetup.frx":0E42
            TabIndex        =   37
            Top             =   0
            Width           =   1335
         End
      End
      Begin VB.PictureBox picHyperlink 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   3
         Left            =   345
         ScaleHeight     =   255
         ScaleWidth      =   1575
         TabIndex        =   20
         Top             =   1200
         Width           =   1575
         Begin VB.Timer tmrHyperlink 
            Index           =   3
            Left            =   0
            Top             =   120
         End
         Begin VB.Label lblHyperlink 
            AutoSize        =   -1  'True
            Caption         =   "Registerinstellingen"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   3
            Left            =   0
            MouseIcon       =   "frmSetup.frx":128C
            TabIndex        =   21
            Top             =   0
            Width           =   1395
         End
      End
      Begin VB.PictureBox picHyperlink 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   1
         Left            =   345
         ScaleHeight     =   255
         ScaleWidth      =   1575
         TabIndex        =   17
         Top             =   480
         Width           =   1575
         Begin VB.Timer tmrHyperlink 
            Index           =   1
            Left            =   0
            Top             =   120
         End
         Begin VB.Label lblHyperlink 
            AutoSize        =   -1  'True
            Caption         =   "Notesbestand"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   0
            MouseIcon       =   "frmSetup.frx":16D6
            TabIndex        =   18
            Top             =   0
            Width           =   1005
         End
      End
      Begin VB.PictureBox picHyperlink 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   0
         Left            =   345
         ScaleHeight     =   255
         ScaleWidth      =   1575
         TabIndex        =   3
         Top             =   120
         Width           =   1575
         Begin VB.Timer tmrHyperlink 
            Index           =   0
            Left            =   0
            Top             =   120
         End
         Begin VB.Label lblHyperlink 
            AutoSize        =   -1  'True
            Caption         =   "Gebruikersopties"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   0
            MouseIcon       =   "frmSetup.frx":1B20
            TabIndex        =   4
            Top             =   0
            Width           =   1200
         End
      End
      Begin VB.Label lblHyperlinkSelected 
         BackStyle       =   0  'Transparent
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   135
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   5400
      TabIndex        =   1
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton cmdAnnuleren 
      Caption         =   "Annuleren"
      Height          =   375
      Left            =   6600
      TabIndex        =   0
      Top             =   3360
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog cdlFontDefault 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame framePanel 
      BorderStyle     =   0  'None
      Height          =   3015
      Index           =   0
      Left            =   2520
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   5175
      Begin VB.CommandButton cmdFontDefaultEdit 
         Caption         =   "Wijzigen..."
         Height          =   375
         Left            =   4080
         TabIndex        =   24
         Top             =   2640
         Width           =   1095
      End
      Begin VB.PictureBox picFontDefault 
         BackColor       =   &H80000005&
         Height          =   495
         Left            =   1920
         ScaleHeight     =   435
         ScaleWidth      =   3195
         TabIndex        =   22
         Top             =   2040
         Width           =   3255
         Begin VB.Label lblFontDefault 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "naam grootte"
            Height          =   195
            Left            =   1125
            TabIndex        =   23
            Top             =   0
            Width           =   945
         End
      End
      Begin VB.CheckBox chkExpandRootNode 
         Caption         =   "Klap het eerste item in de lijst bij het opstarten uit"
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
         Left            =   270
         TabIndex        =   11
         Top             =   360
         Width           =   3975
      End
      Begin VB.CheckBox chkAutoDeleteGroep 
         Caption         =   "Verwijder een groep automatisch als alle items verwijderd zijn"
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
         Left            =   270
         TabIndex        =   10
         Top             =   600
         Width           =   4935
      End
      Begin VB.CheckBox chkAskSaveAfterDelete 
         Caption         =   "Vraag of de notes opgeslagen moeten worden"
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
         Left            =   270
         TabIndex        =   9
         Top             =   1320
         Width           =   6015
      End
      Begin VB.CheckBox chkAskSaveAfterNew 
         Caption         =   "Vraag of de notes opgeslagen moeten worden"
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
         Left            =   270
         TabIndex        =   8
         Top             =   840
         Width           =   3855
      End
      Begin VB.Line Line6 
         BorderColor     =   &H80000014&
         X1              =   270
         X2              =   5160
         Y1              =   1905
         Y2              =   1905
      End
      Begin VB.Line Line5 
         BorderColor     =   &H80000010&
         X1              =   270
         X2              =   5160
         Y1              =   1890
         Y2              =   1890
      End
      Begin VB.Label Label10 
         Caption         =   "na het verwijderen van een item"
         Height          =   255
         Left            =   600
         TabIndex        =   39
         Top             =   1560
         Width           =   2415
      End
      Begin VB.Label Label9 
         Caption         =   "na het maken van een item"
         Height          =   255
         Left            =   600
         TabIndex        =   38
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "Standaard lettertype"
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
         Left            =   270
         TabIndex        =   25
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Gebruikersopties"
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
         TabIndex        =   7
         Top             =   0
         Width           =   1305
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000014&
         Index           =   0
         X1              =   0
         X2              =   6390
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000010&
         Index           =   0
         X1              =   0
         X2              =   6390
         Y1              =   105
         Y2              =   105
      End
   End
   Begin VB.Frame framePanel 
      BorderStyle     =   0  'None
      Height          =   3015
      Index           =   1
      Left            =   2520
      TabIndex        =   12
      Top             =   120
      Visible         =   0   'False
      Width           =   5175
      Begin VB.TextBox txtNotesFileDefault 
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
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   360
         Width           =   3015
      End
      Begin VB.CommandButton cmdNotesFileDefaultBrowse 
         Caption         =   "Bladeren..."
         Height          =   375
         Left            =   4080
         TabIndex        =   15
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton cmdNotesFileDefaultNotesFileOpen 
         Caption         =   "Geopend notesbestand"
         Height          =   375
         Left            =   2160
         TabIndex        =   14
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label5 
         Caption         =   "Standaard notesbestand"
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
         Left            =   270
         TabIndex        =   19
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Notesbestand"
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
         TabIndex        =   13
         Top             =   0
         Width           =   1110
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000010&
         Index           =   1
         X1              =   0
         X2              =   6390
         Y1              =   105
         Y2              =   105
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000014&
         Index           =   1
         X1              =   0
         X2              =   6390
         Y1              =   120
         Y2              =   120
      End
   End
   Begin VB.Frame framePanel 
      BorderStyle     =   0  'None
      Height          =   3015
      Index           =   2
      Left            =   2520
      TabIndex        =   26
      Top             =   120
      Visible         =   0   'False
      Width           =   5175
      Begin VB.CommandButton cmdRegistryContextDelete 
         Caption         =   "Verwijderen"
         Height          =   375
         Left            =   4080
         TabIndex        =   29
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CheckBox chkRegistryContextCreate 
         Caption         =   "Maak de bestandskoppeling en de contextmenus"
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
         TabIndex        =   28
         Top             =   1080
         Width           =   3855
      End
      Begin VB.Label Label6 
         Caption         =   "Bestandskoppeling"
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
         TabIndex        =   27
         Top             =   0
         Width           =   1440
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000014&
         Index           =   2
         X1              =   0
         X2              =   6390
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000010&
         Index           =   2
         X1              =   0
         X2              =   6390
         Y1              =   105
         Y2              =   105
      End
      Begin VB.Label Label8 
         Caption         =   "Door te dubbelklikken op een notesbestand of door met de rechtermuisknop op een notesbestand te klikken kan je het direct openen."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   0
         TabIndex        =   30
         Top             =   360
         Width           =   5175
      End
   End
   Begin VB.Frame framePanel 
      BorderStyle     =   0  'None
      Height          =   3015
      Index           =   3
      Left            =   2520
      TabIndex        =   31
      Top             =   120
      Visible         =   0   'False
      Width           =   5175
      Begin VB.CheckBox chkRegistryPositionSave 
         Caption         =   "Vensterpositie opslaan bij afsluiten"
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
         TabIndex        =   40
         Top             =   960
         Width           =   2895
      End
      Begin VB.CommandButton cmdRegistryDelete 
         Caption         =   "Verwijderen"
         Height          =   375
         Left            =   4080
         TabIndex        =   34
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CheckBox chkRegistrySetupSave 
         Caption         =   "Instellingen opslaan"
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
         TabIndex        =   33
         Top             =   1200
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "WMS Notes bewaart alle instellingen van het programma en de afmeting en positie van het hoofdvenster in het register."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   35
         Top             =   360
         Width           =   5175
      End
      Begin VB.Label Label7 
         Caption         =   "Registerinstellingen"
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
         TabIndex        =   32
         Top             =   0
         Width           =   1500
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000010&
         Index           =   3
         X1              =   0
         X2              =   6390
         Y1              =   105
         Y2              =   105
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000014&
         Index           =   3
         X1              =   0
         X2              =   6390
         Y1              =   120
         Y2              =   120
      End
   End
   Begin VB.Line Line11 
      BorderColor     =   &H80000010&
      X1              =   120
      X2              =   7680
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line12 
      BorderColor     =   &H80000014&
      X1              =   120
      X2              =   7680
      Y1              =   3255
      Y2              =   3255
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   -15
      X2              =   10000
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      X1              =   -15
      X2              =   10000
      Y1              =   15
      Y2              =   15
   End
End
Attribute VB_Name = "frmSetup"
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

'Het aantal panels op deze Form.
Private Const cntPanels As Integer = 4

'Dit houdt bij of de muis zich boven een hyperlink bevindt.
Private MouseOver(0 To cntPanels - 1) As Boolean

'Dit zijn de kleuren die de hyperlink kan krijgen.
Private Const clrHyperlink = vbBlack
Private Const clrHyperlinkMouseOver = 11757824 '62713
Private Const clrHyperlinkMouseDown = 11757824 '62713
Private Const clrHyperlinkBack = vbWhite
Private Sub UpdateHyperlinkSize(ByVal Index As Integer)
    'Geef picHyperlink dezelfde afmetingen als lblHyperlink,
    'zodat picHyperlink niet opvalt.
    picHyperlink(Index).Width = lblHyperlink(Index).Width
    picHyperlink(Index).Height = lblHyperlink(Index).Height
End Sub
Private Sub cmdAnnuleren_Click()
    Canceled = True
    Me.Hide
End Sub
Private Sub cmdFontDefaultEdit_Click()
    On Error GoTo mnuFontDefaultEditError

    'Init: cdlFontDefault
    With cdlFontDefault
        .CancelError = True
        .Flags = cdlCFBoth Or cdlCFEffects
        .FontName = lblFontDefault.Font.Name
        .FontSize = lblFontDefault.Font.Size
        .FontBold = lblFontDefault.Font.Bold
        .FontItalic = lblFontDefault.Font.Italic
        .FontUnderline = lblFontDefault.Font.Underline
        .FontStrikethru = lblFontDefault.Font.Strikethrough
        .Color = lblFontDefault.ForeColor
    End With

    'Show: Font Dialog
    cdlFontDefault.ShowFont

    'Update: lblFontDefault
    With cdlFontDefault
        lblFontDefault.Font.Name = .FontName
        lblFontDefault.Font.Size = .FontSize
        lblFontDefault.Font.Bold = .FontBold
        lblFontDefault.Font.Italic = .FontItalic
        lblFontDefault.Font.Underline = .FontUnderline
        lblFontDefault.Font.Strikethrough = .FontStrikethru
        lblFontDefault.ForeColor = .Color
    End With
    lblFontDefault.Caption = lblFontDefault.Font.Name & " " & CInt(lblFontDefault.Font.Size) & " pt"

    Exit Sub
mnuFontDefaultEditError:
    Select Case Err.Number
        Case 0, cdlCancel
            Err.Clear
        Case Else
            Debug.Print Err.Description
    End Select
End Sub

Private Sub cmdNotesFileDefaultBrowse_Click()
    Dim File As String

    If SelectFileOpen(File) Then txtNotesFileDefault.Text = File
End Sub
Private Sub cmdNotesFileDefaultNotesFileOpen_Click()
    If Not (NotesFileOpen = "") Then txtNotesFileDefault.Text = NotesFileOpen
End Sub
Private Sub cmdOK_Click()
    Canceled = False
    Me.Hide
End Sub
Private Sub cmdRegistryContextDelete_Click()
    RegistryContextDelete
End Sub
Private Sub cmdRegistryDelete_Click()
    RegistryDelete
End Sub
Private Sub Form_Load()
    Dim iHyperlink As Integer

    'Init: Controls
    ' Hyperlinks
    For iHyperlink = 0 To cntPanels - 1
        lblHyperlink(iHyperlink).ForeColor = clrHyperlink
        lblHyperlink(iHyperlink).BackColor = clrHyperlinkBack
'        picHyperlink(iHyperlink).BackColor = clrHyperlinkBack

        UpdateHyperlinkSize iHyperlink
    Next
    lblHyperlink_Click (0)
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Cancel = 1
        cmdAnnuleren_Click
    End If
End Sub
Private Sub lblHyperlink_Click(Index As Integer)
    Dim iPanel As Integer

    'Update: lblHyperlinkSelected
    lblHyperlinkSelected.Top = picHyperlink(Index).Top

    'Update: framePanel
    For iPanel = 0 To cntPanels - 1
        If framePanel(iPanel).Visible Then framePanel(iPanel).Visible = False
    Next iPanel
    framePanel(Index).Visible = True
End Sub
Private Sub lblHyperlink_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If MouseOver(Index) Then
        'Update: lblHyperlink
        lblHyperlink(Index).ForeColor = clrHyperlinkMouseDown
        lblHyperlink(Index).Refresh
    End If
End Sub
Private Sub lblHyperlink_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If MouseOver(Index) Then
        'Update: lblHyperlink
        lblHyperlink(Index).ForeColor = clrHyperlinkMouseOver
        lblHyperlink(Index).Refresh
    End If
End Sub
Private Sub lblHyperlink_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Controleer of de cursor zich wel boven lblHyperlink bevindt.
    If (X < 0) Or (X > picHyperlink(Index).Width) Or (Y < 0) Or (Y > picHyperlink(Index).Height) Then Exit Sub

    If Not MouseOver(Index) Then
        'Update: MouseOver()
        MouseOver(Index) = True

        'Init: trmHyperlink
        tmrHyperlink(Index).Interval = 10
        tmrHyperlink(Index).Enabled = True

        'Update: lblHyperlink
        lblHyperlink(Index).ForeColor = clrHyperlinkMouseOver
'        lblHyperlink(Index).Font.Underline = True
    End If
End Sub
Private Sub tmrHyperlink_Timer(Index As Integer)
    Dim Point As POINTAPI
    Dim X As Long
    Dim Y As Long

    'Bepaal de positie van de cursor.
    GetCursorPos Point
    ScreenToClient picHyperlink(Index).hWnd, Point

    'Converteer naar Twips.
    X = Point.X * Screen.TwipsPerPixelX
    Y = Point.Y * Screen.TwipsPerPixelY

    If (X < 0) Or (X > picHyperlink(Index).Width) Or (Y < 0) Or (Y > picHyperlink(Index).Height) Then
        'De cursor bevindt zich buiten picHyperlink, dus reset alles.

        'Update: MouseOver()
        MouseOver(Index) = False

        'Update: tmrHyperlink
        tmrHyperlink(Index).Enabled = False

        'Update: lblHyperlink
        lblHyperlink(Index).ForeColor = clrHyperlink
'        lblHyperlink(Index).Font.Underline = False
    End If
End Sub
Private Sub txtNotesFileDefault_KeyPress(KeyAscii As Integer)
    Beep
End Sub
