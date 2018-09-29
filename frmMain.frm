VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "WMS Notes"
   ClientHeight    =   5310
   ClientLeft      =   -15
   ClientTop       =   735
   ClientWidth     =   8880
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5310
   ScaleWidth      =   8880
   Begin VB.TextBox txtBeschrijving 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2760
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   875
      Width           =   3165
   End
   Begin VB.TextBox txtTitel 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   480
      Width           =   3165
   End
   Begin MSComctlLib.Toolbar tbrToolbar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlToolbar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "FileSave"
            Object.ToolTipText     =   "Opslaan (Ctrl+S)"
            ImageKey        =   "FileSave"
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "FileReload"
            Object.ToolTipText     =   "Herladen (F5)"
            ImageKey        =   "FileReload"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "FileNoteNew"
            Object.ToolTipText     =   "Nieuwe note (Shift+Ins)"
            ImageKey        =   "FileNoteNew"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "FileNoteDelete"
            Object.ToolTipText     =   "Note verwijderen (Shift+Del)"
            ImageKey        =   "FileNoteDelete"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "FileGroepNew"
            Object.ToolTipText     =   "Nieuwe groep"
            ImageKey        =   "FileGroepNew"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "FileGroepDelete"
            Object.ToolTipText     =   "Groep verwijderen"
            ImageKey        =   "FileGroepDelete"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "EditZoeken"
            Object.ToolTipText     =   "Zoeken... (Ctrl+F)"
            ImageKey        =   "EditZoeken"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "NoteInfo"
            ImageKey        =   "NoteInfo"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "NoteEditProperties"
            Object.ToolTipText     =   "Trefwoorden wijzigen... (Ctrl+P)"
            ImageKey        =   "NoteEditProperties"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
      Begin VB.CheckBox chkIsEncrypted 
         Caption         =   "Versleutel notes"
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
         Left            =   3720
         TabIndex        =   7
         Top             =   45
         Width           =   1455
      End
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
         IMEMode         =   3  'DISABLE
         Left            =   5205
         MaxLength       =   1000
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   30
         Width           =   1455
      End
   End
   Begin MSComctlLib.TreeView lstItems 
      Height          =   3615
      Left            =   60
      TabIndex        =   1
      Top             =   480
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   6376
      _Version        =   393217
      Indentation     =   0
      Style           =   7
      HotTracking     =   -1  'True
      ImageList       =   "imlTreeView"
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
   Begin RichTextLib.RichTextBox txtTekst 
      Height          =   2430
      HelpContextID   =   1130
      Left            =   2760
      TabIndex        =   4
      Top             =   1665
      Width           =   3405
      _ExtentX        =   6006
      _ExtentY        =   4286
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      DisableNoScroll =   -1  'True
      BulletIndent    =   284
      TextRTF         =   $"frmMain.frx":0E42
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog cdlFont 
      Left            =   -120
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin MSComctlLib.ImageList imlTreeView 
      Left            =   -120
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   65280
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0EB4
            Key             =   "INT_Desktop"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1408
            Key             =   "INT_GDicht"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1964
            Key             =   "INT_GOpen"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1EB8
            Key             =   "INT_NOpen"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":240C
            Key             =   "INT_NDicht"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlToolbar 
      Left            =   -120
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   65280
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2960
            Key             =   "FileNoteNew"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2EB4
            Key             =   "FileNoteDelete"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3408
            Key             =   "FileGroepNew"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":395C
            Key             =   "FileGroepDelete"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3EB0
            Key             =   "EditZoeken"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4404
            Key             =   "NoteInfo"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":495C
            Key             =   "NoteEditProperties"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4EB0
            Key             =   "FileSave"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5404
            Key             =   "FileReload"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cdlFile 
      Left            =   -120
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblSplitter 
      Height          =   5175
      Left            =   2700
      MousePointer    =   9  'Size W E
      TabIndex        =   5
      Top             =   480
      Width           =   60
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      X1              =   0
      X2              =   19880
      Y1              =   375
      Y2              =   375
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   0
      X2              =   19880
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&Bestand"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&Nieuw notesbestand"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Openen..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileReload 
         Caption         =   "&Herladen"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "O&pslaan"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Ops&laan als..."
      End
      Begin VB.Menu mnuFileImport 
         Caption         =   "I&mporteren..."
      End
      Begin VB.Menu mnuStreep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileNoteNew 
         Caption         =   "&Nieuwe note"
         Shortcut        =   +{INSERT}
      End
      Begin VB.Menu mnuFileNoteDelete 
         Caption         =   "Note &verwijderen"
         Shortcut        =   +{DEL}
      End
      Begin VB.Menu mnuStreep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileGroepNew 
         Caption         =   "Nieuwe &groep"
      End
      Begin VB.Menu mnuFileGroepDelete 
         Caption         =   "&Groep verwijderen"
      End
      Begin VB.Menu mnuStreep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileAfsluiten 
         Caption         =   "&Afsluiten"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Be&werken"
      Begin VB.Menu mnuEditCut 
         Caption         =   "K&nippen"
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Kopiëren"
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Plakken"
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "&Verwijderen"
      End
      Begin VB.Menu mnuStreep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditSelectAll 
         Caption         =   "&Alles selecteren"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuStreep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditZoeken 
         Caption         =   "&Zoeken..."
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuStreep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditSetup 
         Caption         =   "Opties..."
      End
   End
   Begin VB.Menu mnuNote 
      Caption         =   "&Note"
      Begin VB.Menu mnuNoteInfo 
         Caption         =   "Note &informatie..."
      End
      Begin VB.Menu mnuStreep7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNoteEditName 
         Caption         =   "&Naam wijzigen"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuNoteEditProperties 
         Caption         =   "Tref&woorden wijzigen..."
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFont 
         Caption         =   "Opmaa&k"
         Begin VB.Menu mnuFontBold 
            Caption         =   "&Vet"
            Shortcut        =   ^B
         End
         Begin VB.Menu mnuFontItalic 
            Caption         =   "&Cursief"
            Shortcut        =   ^I
         End
         Begin VB.Menu mnuFontUnderline 
            Caption         =   "Onder&strepen"
            Shortcut        =   ^U
         End
         Begin VB.Menu mnuStreep8 
            Caption         =   "-"
         End
         Begin VB.Menu mnuFontAlignLeft 
            Caption         =   "Links &uitlijnen"
            Shortcut        =   ^L
         End
         Begin VB.Menu mnuFontAlignCenter 
            Caption         =   "Cen&treren"
            Shortcut        =   ^E
         End
         Begin VB.Menu mnuFontAlignRight 
            Caption         =   "Rechts &uitlijnen"
            Shortcut        =   ^R
         End
         Begin VB.Menu mnuStreep9 
            Caption         =   "-"
         End
         Begin VB.Menu mnuFontBullet 
            Caption         =   "&Opsommingstekens"
         End
         Begin VB.Menu mnuStreep10 
            Caption         =   "-"
         End
         Begin VB.Menu mnuFontArial 
            Caption         =   "&Lettertype: Arial"
         End
         Begin VB.Menu mnuFontCourierNew 
            Caption         =   "&Lettertype: Courier New"
         End
         Begin VB.Menu mnuFontTahoma 
            Caption         =   "&Lettertype: Tahoma"
         End
         Begin VB.Menu mnuStreep11 
            Caption         =   "-"
         End
         Begin VB.Menu mnuFontEdit 
            Caption         =   "&Lettertype kiezen..."
            Shortcut        =   ^D
         End
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpInhoud 
         Caption         =   "Inhoudsopgave en inde&x"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuStreep12 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInfo 
         Caption         =   "&Informatie..."
      End
   End
   Begin VB.Menu mnuPopup1 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuPopup1NoteEditName 
         Caption         =   "&Naam wijzigen"
      End
      Begin VB.Menu mnuPopup1NoteEditProperties 
         Caption         =   "Tref&woorden wijzigen..."
      End
      Begin VB.Menu mnuStreep14 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopup1FileNoteDelete 
         Caption         =   "Note &verwijderen"
      End
      Begin VB.Menu mnuPopup1FileGroepDelete 
         Caption         =   "&Groep verwijderen"
      End
   End
   Begin VB.Menu mnuPopup2 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuPopup2EditCut 
         Caption         =   "K&nippen"
      End
      Begin VB.Menu mnuPopup2EditCopy 
         Caption         =   "&Kopiëren"
      End
      Begin VB.Menu mnuPopup2EditPaste 
         Caption         =   "&Plakken"
      End
      Begin VB.Menu mnuStreep15 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopup2FontBold 
         Caption         =   "&Vet"
      End
      Begin VB.Menu mnuPopup2FontItalic 
         Caption         =   "&Cursief"
      End
      Begin VB.Menu mnuPopup2FontUnderline 
         Caption         =   "Onder&strepen"
      End
      Begin VB.Menu mnuStreep16 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopup2FontAlignLeft 
         Caption         =   "Links &uitlijnen"
      End
      Begin VB.Menu mnuPopup2FontAlignCenter 
         Caption         =   "Cen&treren"
      End
      Begin VB.Menu mnuPopup2FontAlignRight 
         Caption         =   "Rechts &uitlijnen"
      End
      Begin VB.Menu mnuStreep17 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopup2FontBullet 
         Caption         =   "&Opsommingstekens"
      End
      Begin VB.Menu mnuStreep18 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopup2FontArial 
         Caption         =   "&Lettertype: Arial"
      End
      Begin VB.Menu mnuPopup2FontCourierNew 
         Caption         =   "&Lettertype: Courier New"
      End
      Begin VB.Menu mnuPopup2FontTahoma 
         Caption         =   "&Lettertype: Tahoma"
      End
      Begin VB.Menu mnuStreep19 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopup2FontEdit 
         Caption         =   "&Lettertype kiezen..."
      End
   End
End
Attribute VB_Name = "frmMain"
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

'Drag-And-Drop Operation
Private NodeMove As Node             'De te verplaatsen node
Private DragDropOperation As Boolean 'Drag-And-Drop Operation bezig?

'Splitter
Private SplitterMoving As Boolean    'Geeft aan of de Splitter versleept wordt
Sub SizeControls(ByVal SplitterLeft As Integer)
    On Error Resume Next
    Dim TekstLeft As Integer, TekstWidth As Integer

    'Check: SplitterLeft
    If SplitterLeft - lstItems.Left < 450 Then SplitterLeft = 510
    If frmMain.ScaleWidth - SplitterLeft < 650 Then SplitterLeft = frmMain.ScaleWidth - 650

    'Update: Top; Height
    If frmMain.ScaleHeight - (lstItems.Top + 60) > 0 Then _
        lstItems.Height = frmMain.ScaleHeight - (lstItems.Top + 60)
    If frmMain.ScaleHeight - (lblSplitter.Top + 60) > 0 Then _
        lblSplitter.Height = frmMain.ScaleHeight - (lblSplitter.Top + 60)
    If frmMain.ScaleHeight - (txtTekst.Top + 60) > 0 Then _
        txtTekst.Height = frmMain.ScaleHeight - (txtTekst.Top + 60)

    'Update: Left; Width
    lblSplitter.Left = SplitterLeft
    lstItems.Width = SplitterLeft - lstItems.Left

    TekstLeft = SplitterLeft + 60
    TekstWidth = frmMain.ScaleWidth - (txtTekst.Left + 60)
    txtTitel.Left = TekstLeft
    txtTitel.Width = TekstWidth
    txtBeschrijving.Left = TekstLeft
    txtBeschrijving.Width = TekstWidth
    txtTekst.Left = TekstLeft
    txtTekst.Width = TekstWidth
End Sub
Public Sub UpdatemnuEdit()
    'Update: mnuEdit
    If Not (Clipboard.GetText(vbCFText) = "") Or Not (Clipboard.GetText(vbCFRTF) = "") Then
        mnuEditPaste.Enabled = True
    Else
        mnuEditPaste.Enabled = False
    End If
    If Not (txtBeschrijving.SelLength = 0) Or Not (frmMain.txtTekst.SelLength = 0) Then
        mnuEditCut.Enabled = True
        mnuEditCopy.Enabled = True
        mnuEditDelete.Enabled = True
    Else
        mnuEditCut.Enabled = False
        mnuEditCopy.Enabled = False
        mnuEditDelete.Enabled = False
    End If
End Sub
Public Sub UpdatemnuFont()
    'Update: mnuFont
    'Let op: Er kan Null als resultaat gegeven worden!

    If txtTekst.SelBold = True Then _
        mnuFontBold.Checked = True Else _
        mnuFontBold.Checked = False
    If txtTekst.SelItalic = True Then _
        mnuFontItalic.Checked = True Else _
        mnuFontItalic.Checked = False
    If txtTekst.SelUnderline = True Then _
        mnuFontUnderline.Checked = True Else _
        mnuFontUnderline.Checked = False

    Select Case txtTekst.SelAlignment
        Case Null
            mnuFontAlignLeft.Checked = False
            mnuFontAlignCenter.Checked = False
            mnuFontAlignRight.Checked = False
        Case rtfLeft
            mnuFontAlignLeft.Checked = True
            mnuFontAlignCenter.Checked = False
            mnuFontAlignRight.Checked = False
        Case rtfCenter
            mnuFontAlignLeft.Checked = False
            mnuFontAlignCenter.Checked = True
            mnuFontAlignRight.Checked = False
        Case rtfRight
            mnuFontAlignLeft.Checked = False
            mnuFontAlignCenter.Checked = False
            mnuFontAlignRight.Checked = True
    End Select

    If txtTekst.SelBullet = True Then _
        mnuFontBullet.Checked = True Else _
        mnuFontBullet.Checked = False

    If txtTekst.SelFontName = "Arial" Then _
        mnuFontArial.Checked = True Else _
        mnuFontArial.Checked = False
    If txtTekst.SelFontName = "Courier New" Then _
        mnuFontCourierNew.Checked = True Else _
        mnuFontCourierNew.Checked = False
    If txtTekst.SelFontName = "Tahoma" Then _
        mnuFontTahoma.Checked = True Else _
        mnuFontTahoma.Checked = False
End Sub
Private Sub chkIsEncrypted_Click()
    If chkIsEncrypted.Value = vbChecked Then _
        IsEncrypted = True Else _
        IsEncrypted = False
    txtPassword.Enabled = IsEncrypted
End Sub
Private Sub Form_Load()
    'Init: mnuEdit
    mnuEditCut.Caption = "K&nippen" & vbTab & "Ctrl+X"
    mnuEditCopy.Caption = "&Kopiëren" & vbTab & "Ctrl+C"
    mnuEditPaste.Caption = "&Plakken" & vbTab & "Ctrl+V"
    mnuEditDelete.Caption = "&Verwijderen" & vbTab & "Del"

    'Init: mnuPopup1
    mnuPopup1NoteEditName.Caption = "&Naam wijzigen" & vbTab & "F2"
    mnuPopup1NoteEditProperties.Caption = "Tref&woorden wijzigen..." & vbTab & "Ctrl+P"
    mnuPopup1FileNoteDelete.Caption = "Note &verwijderen" & vbTab & "Shift+Del"
    mnuPopup1FileGroepDelete.Caption = "&Groep verwijderen"

    'Init: mnuPopup2
    mnuPopup2EditCut.Caption = "K&nippen" & vbTab & "Ctrl+X"
    mnuPopup2EditCopy.Caption = "&Kopiëren" & vbTab & "Ctrl+C"
    mnuPopup2EditPaste.Caption = "&Plakken" & vbTab & "Ctrl+V"
    mnuPopup2FontBold.Caption = "&Vet" & vbTab & "Ctrl+B"
    mnuPopup2FontItalic.Caption = "&Cursief" & vbTab & "Ctrl+I"
    mnuPopup2FontUnderline.Caption = "Onder&strepen" & vbTab & "Ctrl+U"
    mnuPopup2FontAlignLeft.Caption = "Links &uitlijnen" & vbTab & "Ctrl+L"
    mnuPopup2FontAlignCenter.Caption = "Cen&treren" & vbTab & "Ctrl+E"
    mnuPopup2FontAlignRight.Caption = "Rechts &uitlijnen" & vbTab & "Ctrl+R"
    mnuPopup2FontEdit.Caption = "&Lettertype kiezen..." & vbTab & "Ctrl+D"
End Sub
Private Sub mnuFileNew_Click()
    'Check: MagInterfere
    If MagInterfere = False Then Exit Sub

    NotesFileNew
End Sub
Private Sub mnuFileOpen_Click()
    Dim File As String

    'Check: MagInterfere
    If MagInterfere = False Then Exit Sub

    If SelectFileOpen(File) Then NotesFileLoad File
End Sub
Private Sub mnuPopup1FileGroepDelete_Click()
    mnuFileGroepDelete_Click
End Sub
Private Sub mnuPopup1FileNoteDelete_Click()
    mnuFileNoteDelete_Click
End Sub
Private Sub mnuPopup1NoteEditName_Click()
    mnuNoteEditName_Click
End Sub
Private Sub mnuPopup1NoteEditProperties_Click()
    mnuNoteEditProperties_Click
End Sub
Private Sub mnuPopup2EditCopy_Click()
    mnuEditCopy_Click
End Sub
Private Sub mnuPopup2EditCut_Click()
    mnuEditCut_Click
End Sub
Private Sub mnuPopup2EditPaste_Click()
    mnuEditPaste_Click
End Sub
Private Sub mnuPopup2FontAlignCenter_Click()
    mnuFontAlignCenter_Click
End Sub
Private Sub mnuPopup2FontAlignLeft_Click()
    mnuFontAlignLeft_Click
End Sub
Private Sub mnuPopup2FontAlignRight_Click()
    mnuFontAlignRight_Click
End Sub
Private Sub mnuPopup2FontArial_Click()
    mnuFontArial_Click
End Sub
Private Sub mnuPopup2FontBold_Click()
    mnuFontBold_Click
End Sub
Private Sub mnuPopup2FontBullet_Click()
    mnuFontBullet_Click
End Sub
Private Sub mnuPopup2FontCourierNew_Click()
    mnuFontCourierNew_Click
End Sub
Private Sub mnuPopup2FontEdit_Click()
    mnuFontEdit_Click
End Sub
Private Sub mnuPopup2FontItalic_Click()
    mnuFontItalic_Click
End Sub
Private Sub mnuPopup2FontTahoma_Click()
    mnuFontTahoma_Click
End Sub
Private Sub mnuPopup2FontUnderline_Click()
    mnuFontUnderline_Click
End Sub
Private Sub txtPassword_Change()
    Password = txtPassword.Text
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    'Check: MagInterfere; ActiveControl
    If MagInterfere = False Then Exit Sub
    If Not (ActiveControl.Name = "txtTekst") Then Exit Sub

    'Check: KeyAscii
    Select Case KeyAscii
        'Bold, Italic, Underline
        Case 2, 9, 21
'            KeyAscii = 0

        'Copy, Paste, Cut
        Case 3, 22, 24
            KeyAscii = 0

        'Select All
        Case 1
            KeyAscii = 0
    End Select
End Sub
Private Sub Form_Resize()
    If frmMain.WindowState = 0 Or frmMain.WindowState = 2 Then _
        SizeControls lblSplitter.Left
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If MainEnd = True Then Cancel = 1
End Sub
Private Sub lblSplitter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Init: SplitterMoving
    SplitterMoving = True
End Sub
Private Sub lblSplitter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Check: SplitterMoving
    If SplitterMoving = False Then Exit Sub

    'Update: Controls
    SizeControls lblSplitter.Left + X
End Sub
Private Sub lblSplitter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Update: SplitterMoving
    SplitterMoving = False
End Sub
Private Sub lstItems_AfterLabelEdit(Cancel As Integer, NewString As String)
    Dim NodeEdited As Node

    'Check: MagInterfere
    If MagInterfere = False Then
        Cancel = 1
        Exit Sub
    End If

    'Check: NewString
    If NewString = "" Then
        Cancel = 1
        Exit Sub
    End If
    If Not (InStr(1, NewString, "\") = 0) Or Not (InStr(1, NewString, "/") = 0) Then
        MsgBox "Een naam van een note of notegroep mag deze tekens niet bevatten: \  /", vbExclamation
        Cancel = 1
        Exit Sub
    End If

    'Update: Note
    Set NodeEdited = lstItems.SelectedItem
    If Left(NodeEdited.Key, 3) = KeyPrefixNote Then
        'Update: Notes()
        WijzigNote "Titel", NewString
        NodeEdited.Parent.Sorted = True

    ElseIf Left(NodeEdited.Key, 3) = KeyPrefixGroep Then
        'Update: Groep
        GroepWijzigNaam NewString
        Cancel = 1
    End If
End Sub
Private Sub lstItems_BeforeLabelEdit(Cancel As Integer)
    'Check: MagInterfere
    If MagInterfere = False Then
        Cancel = 1
        Exit Sub
    End If

    If lstItems.SelectedItem.Key = KeyRoot Then Cancel = 1
End Sub
Private Sub lstItems_DragDrop(Source As Control, X As Single, Y As Single)
    Dim NodeNieuw As Node
    Dim NodeOud As Node

    'Check: MagInterfere
    If MagInterfere = False Then Exit Sub

    'Init: NodeNieuw; NodeOud
    Set NodeNieuw = lstItems.DropHighlight
    Set NodeOud = NodeMove

    'Beëindig de Drag-and-Drop-operation.
    Set NodeMove = Nothing
    Set lstItems.DropHighlight = Nothing
    DragDropOperation = False

    'Check: NodeNieuw
    If NodeNieuw Is Nothing Then Exit Sub

    'Update: NodeNieuw
    If Left(NodeNieuw.Key, 3) = KeyPrefixNote Then
        WijzigNote "Groep", ConvertKeyToGroep(NodeNieuw.Parent.Key)
    ElseIf Left(NodeNieuw.Key, 3) = KeyPrefixGroep Then
        WijzigNote "Groep", ConvertKeyToGroep(NodeNieuw.Key)
    End If

    'Selecteer een nieuwe note en maak hem zichtbaar.
    NodeChange frmMain.lstItems.SelectedItem
End Sub
Private Sub lstItems_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    Dim NodeHighlight As Node

    'Check: MagInterfere; DragDropOperation
    If MagInterfere = False Then Exit Sub
    If DragDropOperation = False Then Exit Sub

    'Init: NodeHighlight
    Set NodeHighlight = lstItems.HitTest(X, Y)
    If NodeHighlight Is Nothing Then Exit Sub
    If NodeHighlight.Key = KeyRoot Then Exit Sub

    'Update: lstItems
    Set lstItems.DropHighlight = NodeHighlight
End Sub
Private Sub lstItems_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim NodeSelected As Node

    'Check: MagInterfere; Button
    If MagInterfere = False Then Exit Sub
    If Not (Button = vbLeftButton) Then Exit Sub

    'Init: NodeSelected
    Set NodeSelected = lstItems.HitTest(X, Y)
    If NodeSelected Is Nothing Then Exit Sub
    If Not (Left(NodeSelected.Key, 3) = KeyPrefixNote) Then Exit Sub

    'Init: NodeMove
    Set NodeMove = NodeSelected

    'Update: lstItems
    Set lstItems.DropHighlight = NodeSelected
End Sub
Private Sub lstItems_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Check: MagInterfere
    If MagInterfere = False Then Exit Sub

    'Controleer de juiste muisknop.
    If Not (Button = vbLeftButton) Then
        'Stop met een eventuele Drag-and-Drop-operation.
        Set NodeMove = Nothing
        Set lstItems.DropHighlight = Nothing
        DragDropOperation = False
        Exit Sub
    End If

    'Init.
    Dim NodeSelected As Node
    Set NodeSelected = lstItems.HitTest(X, Y)

    'Controleer de node.
    If NodeSelected Is Nothing Then Exit Sub
    If Not (Left(NodeSelected.Key, 3) = KeyPrefixNote) Then Exit Sub
    If NodeMove Is Nothing Then Exit Sub
    NodeChange NodeMove

    'Begin de Drag-and-Drop-operation.
    lstItems.DragIcon = lstItems.HitTest(X, Y).CreateDragImage
    lstItems.Drag vbBeginDrag
    DragDropOperation = True
End Sub
Private Sub lstItems_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim NodeClicked As Node

    'Check: MagInterfere; Button
    If MagInterfere = False Then Exit Sub
    If Not (Button = vbRightButton) Then Exit Sub

    'Init: NodeClicked
    Set NodeClicked = lstItems.HitTest(X, Y)

    If (NodeClicked Is Nothing) Or (NodeClicked Is lstItems.Nodes(1)) Then
        'Update: mnuFile
        mnuFileNew.Visible = False
        mnuFileSaveAs.Visible = False
        mnuFileImport.Visible = False
        mnuFileAfsluiten.Visible = False
        mnuFileNoteDelete.Visible = False
        mnuFileGroepDelete.Visible = False
        mnuStreep2.Visible = False
        mnuStreep3.Visible = False

        'Show: mnuFile
        PopupMenu mnuFile

        'Update: mnuFile
        mnuFileImport.Visible = True
        mnuFileAfsluiten.Visible = True
        mnuFileNoteDelete.Visible = True
        mnuFileGroepDelete.Visible = True
        mnuStreep2.Visible = True
        mnuStreep3.Visible = True

    Else
        'Init: mnuPopup1
        mnuPopup1NoteEditName.Enabled = mnuNoteEditName.Enabled
        mnuPopup1NoteEditProperties.Visible = mnuNoteEditProperties.Enabled
        mnuPopup1FileNoteDelete.Visible = mnuFileNoteDelete.Enabled
        mnuPopup1FileGroepDelete.Visible = mnuFileGroepDelete.Enabled

        'Show: mnuPopup1
        PopupMenu mnuPopup1
    End If
End Sub
Private Sub lstItems_NodeClick(ByVal Node As MSComctlLib.Node)
    'Check: MagInterfere
    If MagInterfere = False Then Exit Sub

    NodeChange Node
End Sub
Private Sub mnuFontAlignCenter_Click()
    'Check: MagInterfere; ActiveControl
    If MagInterfere = False Then Exit Sub
    If Not (frmMain.ActiveControl.Name = "txtTekst") Then Exit Sub

    txtTekst.SelAlignment = rtfCenter
End Sub
Private Sub mnuFileAfsluiten_Click()
    'Check: MagInterfere
    If MagInterfere = False Then Exit Sub

    MainEnd
End Sub
Private Sub mnuFontAlignLeft_Click()
    'Check: MagInterfere; ActiveControl
    If MagInterfere = False Then Exit Sub
    If Not (frmMain.ActiveControl.Name = "txtTekst") Then Exit Sub

    txtTekst.SelAlignment = rtfLeft
End Sub
Private Sub mnuFontAlignRight_Click()
    'Check: MagInterfere; ActiveControl
    If MagInterfere = False Then Exit Sub
    If Not (frmMain.ActiveControl.Name = "txtTekst") Then Exit Sub

    txtTekst.SelAlignment = rtfRight
End Sub
Private Sub mnuEdit_Click()
    UpdatemnuEdit
End Sub
Private Sub mnuFontBullet_Click()
    'Check: MagInterfere; ActiveControl
    If MagInterfere = False Then Exit Sub
    If Not (frmMain.ActiveControl.Name = "txtTekst") Then Exit Sub

    txtTekst.SelBullet = Not txtTekst.SelBullet
End Sub
Private Sub mnuEditCopy_Click()
    'Check: MagInterfere
    If MagInterfere = False Then Exit Sub

    'Update: Clipboard
    Select Case ActiveControl.Name
        Case "txtTitel", "txtBeschrijving"
            If Not (ActiveControl.SelLength = 0) Then
                Clipboard.SetText ActiveControl.SelText, vbCFText
                Clipboard.SetText ActiveControl.SelText, vbCFRTF
            End If

        Case "txtTekst"
            If Not (ActiveControl.SelLength = 0) Then
                Clipboard.SetText ActiveControl.SelText, vbCFText
                Clipboard.SetText ActiveControl.SelRTF, vbCFRTF
            End If
    End Select
End Sub
Private Sub mnuEditCut_Click()
    'Check: MagInterfere
    If MagInterfere = False Then Exit Sub

    'Update: Clipboard; ActiveControl
    Select Case ActiveControl.Name
        Case "txtTekst", "txtTitel", "txtBeschrijving"
            If Not (ActiveControl.SelLength = 0) Then
                Clipboard.SetText ActiveControl.SelText, vbCFText
                ActiveControl.SelText = ""
            End If
    End Select
End Sub
Private Sub mnuEditDelete_Click()
    'Check: MagInterfere
    If MagInterfere = False Then Exit Sub

    'Update: ActiveControl
    Select Case ActiveControl.Name
        Case "txtTekst", "txtTitel", "txtBeschrijving"
            ActiveControl.SelText = ""
    End Select
End Sub
Private Sub mnuEditPaste_Click()
    'Check: MagInterfere
    If MagInterfere = False Then Exit Sub

    'Update: ActiveControl
    Select Case ActiveControl.Name
        Case "txtTekst", "txtTitel", "txtBeschrijving"
            If Not ActiveControl.Locked Then _
                ActiveControl.SelText = Clipboard.GetText(vbCFText)
    End Select
End Sub
Private Sub mnuEditSelectAll_Click()
    'Check: MagInterfere
    If MagInterfere = False Then Exit Sub

    'Update: Clipboard
    Select Case ActiveControl.Name
        Case "txtTekst", "txtTitel", "txtBeschrijving"
            ActiveControl.SelStart = 0
            ActiveControl.SelLength = Len(ActiveControl.Text)
    End Select
End Sub
Private Sub mnuFontArial_Click()
    'Check: MagInterfere; ActiveControl
    If MagInterfere = False Then Exit Sub
    If Not (frmMain.ActiveControl.Name = "txtTekst") Then Exit Sub

    txtTekst.SelFontName = "Arial"
End Sub
Private Sub mnuFontBold_Click()
    'Check: MagInterfere; ActiveControl
    If MagInterfere = False Then Exit Sub
    If Not (frmMain.ActiveControl.Name = "txtTekst") Then Exit Sub

    txtTekst.SelBold = Not txtTekst.SelBold
End Sub
Private Sub mnuFontCourierNew_Click()
    'Check: MagInterfere; ActiveControl
    If MagInterfere = False Then Exit Sub
    If Not (frmMain.ActiveControl.Name = "txtTekst") Then Exit Sub

    txtTekst.SelFontName = "Courier New"
End Sub
Private Sub mnuFontItalic_Click()
    'Check: MagInterfere; ActiveControl
    If MagInterfere = False Then Exit Sub
    If Not (frmMain.ActiveControl.Name = "txtTekst") Then Exit Sub

    txtTekst.SelItalic = Not txtTekst.SelItalic
End Sub
Private Sub mnuFontEdit_Click()
    On Error GoTo mnuFontEditError

    'Check: MagInterfere
    If MagInterfere = False Then Exit Sub

    'Init: cdlFont
    With cdlFont
        .CancelError = True
        .Flags = cdlCFBoth Or cdlCFEffects
        If Not IsNull(txtTekst.SelFontName) Then .FontName = txtTekst.SelFontName Else .Flags = .Flags Or cdlCFNoFaceSel
        If Not IsNull(txtTekst.SelFontSize) Then .FontSize = txtTekst.SelFontSize
        If Not IsNull(txtTekst.SelColor) Then .Color = txtTekst.SelColor
        If Not IsNull(txtTekst.SelBold) Then .FontBold = txtTekst.SelBold
        If Not IsNull(txtTekst.SelItalic) Then .FontItalic = txtTekst.SelItalic
        If Not IsNull(txtTekst.SelUnderline) Then .FontUnderline = txtTekst.SelUnderline
        If Not IsNull(txtTekst.SelStrikeThru) Then .FontStrikethru = txtTekst.SelStrikeThru
    End With

    'Show: Font Dialog
    cdlFont.ShowFont

    'Update: Controls
    With cdlFont
        If Not (.FontName = "") Then txtTekst.SelFontName = .FontName
        txtTekst.SelFontSize = .FontSize
        txtTekst.SelColor = .Color
        txtTekst.SelBold = .FontBold
        txtTekst.SelItalic = .FontItalic
        txtTekst.SelUnderline = .FontUnderline
        txtTekst.SelStrikeThru = .FontStrikethru
    End With

    Exit Sub
mnuFontEditError:
    Select Case Err.Number
        Case 0, cdlCancel
            Err.Clear
        Case Else
            Debug.Print Err.Description
    End Select
End Sub
Private Sub mnuFontTahoma_Click()
    'Check: MagInterfere; ActiveControl
    If MagInterfere = False Then Exit Sub
    If Not (frmMain.ActiveControl.Name = "txtTekst") Then Exit Sub

    txtTekst.SelFontName = "Tahoma"
End Sub
Private Sub mnuFontUnderline_Click()
    'Check: MagInterfere; ActiveControl
    If MagInterfere = False Then Exit Sub
    If Not (frmMain.ActiveControl.Name = "txtTekst") Then Exit Sub

    txtTekst.SelUnderline = Not txtTekst.SelUnderline
End Sub
Private Sub mnuFileGroepNew_Click()
    'Check: MagInterfere
    If MagInterfere = False Then Exit Sub

    GroepNew
End Sub
Private Sub mnuFileGroepDelete_Click()
    'Check: MagInterfere
    If MagInterfere = False Then Exit Sub

    GroepDelete
End Sub
Private Sub mnuHelpInhoud_Click()
    HTMLHelp.ShowHelp ShowContents
End Sub
Private Sub mnuInfo_Click()
    'Check: MagInterfere
    If MagInterfere = False Then Exit Sub

    frmInfo.Show 1
End Sub
Private Sub mnuFileReload_Click()
    'Check: MagInterfere
    If MagInterfere = False Then Exit Sub

    NotesFileLoad NotesFileOpen
End Sub
Private Sub mnuFileImport_Click()
    Dim File As String

    'Check: MagInterfere
    If MagInterfere = False Then Exit Sub

    If SelectFileOpen(File) Then NotesFileLoad File, False
End Sub
Private Sub mnuFileNoteNew_Click()
    'Check: MagInterfere
    If MagInterfere = False Then Exit Sub

    NoteNieuw
End Sub
Private Sub mnuFileSave_Click()
    'Check: MagInterfere
    If MagInterfere = False Then Exit Sub

    If NotesFileOpen = "" Then
        mnuFileSaveAs_Click
    Else
        NotesFileSave NotesFileOpen, NotesFileOpen & NotesFileBackupExtension
    End If
End Sub
Private Sub mnuFileSaveAs_Click()
    Dim File As String

    'Check: MagInterfere
    If MagInterfere = False Then Exit Sub

    If SelectFileSave(File) Then NotesFileSave File, File & NotesFileBackupExtension
End Sub
Private Sub mnuFileNoteDelete_Click()
    'Check: MagInterfere
    If MagInterfere = False Then Exit Sub

    NoteDelete
End Sub
Private Sub mnuNoteEditName_Click()
    'Check: MagInterfere
    If MagInterfere = False Then Exit Sub

    On Error Resume Next
    lstItems.StartLabelEdit
End Sub
Private Sub mnuFont_Click()
    UpdatemnuFont
End Sub
Private Sub mnuNoteInfo_Click()
    'Check: MagInterfere
    If MagInterfere = False Then Exit Sub

    LaadNoteInfo True
End Sub
Private Sub mnuNoteEditProperties_Click()
    'Check: MagInterfere
    If MagInterfere = False Then Exit Sub

    LaadNoteInfo False
End Sub
Private Sub mnuEditSetup_Click()
    'Check: MagInterfere
    If MagInterfere = False Then Exit Sub

    SetupEdit
End Sub
Private Sub mnuEditZoeken_Click()
    'Check: MagInterfere
    If MagInterfere = False Then Exit Sub

    frmZoeken.Show 1
End Sub
Private Sub tbrToolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "FileReload"
            mnuFileReload_Click
        Case "FileSave"
            mnuFileSave_Click
        Case "FileNoteNew"
            mnuFileNoteNew_Click
        Case "FileNoteDelete"
            mnuFileNoteDelete_Click
        Case "FileGroepNew"
            mnuFileGroepNew_Click
        Case "FileGroepDelete"
            mnuFileGroepDelete_Click

        Case "EditZoeken"
            mnuEditZoeken_Click

        Case "NoteInfo"
            mnuNoteInfo_Click
        Case "NoteEditProperties"
            mnuNoteEditProperties_Click
    End Select
End Sub
Private Sub txtBeschrijving_Change()
    'Check: MagInterfere
    If MagInterfere = False Then Exit Sub

    WijzigNote "Beschrijving", txtBeschrijving.Text
End Sub
Private Sub txtTekst_Change()
    'Check: MagInterfere
    If MagInterfere = False Then Exit Sub

    WijzigNote "Tekst", txtTekst.TextRTF
End Sub
Private Sub txtTekst_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Check: MagInterfere; Button; txtTekst
    If MagInterfere = False Then Exit Sub
    If Not (Button = vbRightButton) Then Exit Sub
    If txtTekst.Locked Then Exit Sub

    'Update: mnuEdit; mnuFont
    UpdatemnuEdit
    UpdatemnuFont

    'Init: mnuPopup2
    mnuPopup2EditCut.Enabled = mnuEditCut.Enabled
    mnuPopup2EditCopy.Enabled = mnuEditCopy.Enabled
    mnuPopup2EditPaste.Enabled = mnuEditPaste.Enabled
    mnuPopup2FontBold.Checked = mnuFontBold.Checked
    mnuPopup2FontItalic.Checked = mnuFontItalic.Checked
    mnuPopup2FontUnderline.Checked = mnuFontUnderline.Checked
    mnuPopup2FontAlignLeft.Checked = mnuFontAlignLeft.Checked
    mnuPopup2FontAlignCenter.Checked = mnuFontAlignCenter.Checked
    mnuPopup2FontAlignRight.Checked = mnuFontAlignRight.Checked
    mnuPopup2FontBullet.Checked = mnuFontBullet.Checked
    mnuPopup2FontArial.Checked = mnuFontArial.Checked
    mnuPopup2FontCourierNew.Checked = mnuFontCourierNew.Checked
    mnuPopup2FontTahoma.Checked = mnuFontTahoma.Checked

    'Show: mnuPopup2
    PopupMenu mnuPopup2
End Sub
