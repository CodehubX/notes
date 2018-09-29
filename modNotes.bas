Attribute VB_Name = "modNotes"
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

'--------------------------------------------------------------------
'Constants
'--------------------------------------------------------------------
Public Const NotesFileCurrentVersion As Integer = 3
Public Const NotesFileNameDefault As String = "Notes"
Public Const NotesFileExtension As String = ".nf" & NotesFileCurrentVersion
Public Const NotesFileExtensionOld1 As String = ".nts"
Public Const NotesFileBackupExtension As String = ".nfb"

Public Const KeyRoot As String = "RT Root"
Public Const KeyPrefixGroep As String = "NG "
Public Const KeyPrefixNote As String = "NE "

'--------------------------------------------------------------------
'Variabelen
'--------------------------------------------------------------------
Public NotesFileDefault As String
Public FontDefaultName As String, FontDefaultSize As Single, FontDefaultBold As Boolean, FontDefaultItalic As Boolean, FontDefaultUnderline As Boolean, FontDefaultStrikeThru As Boolean, FontDefaultColor As Long

Public NotesFileOpen As String
Public IsDirty As Boolean
Public IsEncrypted As Boolean
Public Password As String

Public Notes() As clsNote
Public cntNotes As Integer
Public curNote As Integer
'--------------------------------------------------------------------
'Deze procedure kijkt of er onder een bepaalde groep notes zitten.
'--------------------------------------------------------------------
Public Function GroepHasNotes(ByVal Groep As String) As Boolean
    On Error GoTo GroepHasNotesError
    Dim iNote As Integer

    'Init: GroepHasNotes
    GroepHasNotes = False

    'Check: Notes()
    If cntNotes = 0 Then Exit Function
    For Each Note In Notes()
        If Left(Note.Groep, Len(Groep)) = Groep Then
        If Len(Note.Groep) = Len(Groep) Or Mid(Note.Groep, Len(Groep) + 1, 1) = "\" Then
            GroepHasNotes = True
            Exit For
        End If
        End If
    Next

    Exit Function
GroepHasNotesError:
    ErrorHandling "modNotes", "GroepHasNotes", Err
    Resume Next
End Function
'--------------------------------------------------------------------
'Deze procedure laadt een bepaalde note.
'De procedure mag alleen door NodeChange worden aangeroepen!!!
'--------------------------------------------------------------------
Public Sub LaadNote(ByVal NodeKey As String)
    On Error GoTo LaadNoteError

    'Init: iNote
    iNote = Mid(NodeKey, 4)

    'Update: frmMain
    frmMain.Caption = Programmanaam & " - " & Notes(iNote).Titel

    'Update: Controls
    If Not (frmMain.lstItems.SelectedItem Is Notes(iNote).Node) Then
        Set frmMain.lstItems.SelectedItem = Notes(iNote).Node
    End If
    frmMain.txtTitel.Text = Notes(iNote).Titel
    frmMain.txtTitel.ToolTipText = "[" & Join(Notes(iNote).ZoekTermen, ", ") & "]"
    frmMain.txtBeschrijving.Text = Notes(iNote).Beschrijving
    frmMain.txtTekst.TextRTF = Notes(iNote).Tekst

    Exit Sub
LaadNoteError:
    ErrorHandling "modNotes", "LaadNote", Err, True
    Resume Next
End Sub
'--------------------------------------------------------------------
'Deze procedure verwijdert een groep uit de lijst.
'--------------------------------------------------------------------
Public Sub GroepDelete()
    On Error GoTo GroepDeleteError
    Dim AutoDeleteGroepBackup As Boolean
    Dim GroepDelete As String
    Dim iNote As Integer
    Dim iNoteFree As Integer, iNoteUsed As Integer

    'Update: MagInterfere; MousePointer
    MagInterfere = False
    Screen.MousePointer = vbHourglass

    'Update: AutoDeleteGroep
    AutoDeleteGroepBackup = AutoDeleteGroep
    AutoDeleteGroep = False

    'Init: GroepDelete
    GroepDelete = ConvertKeyToGroep(frmMain.lstItems.SelectedItem.Key)

    'Delete: Groep
    If GroepHasNotes(GroepDelete) Then
        'Weet u het zeker?
        If (MsgBox("Weet u zeker dat u deze groep met en alle notes en subgroepen wilt verwijderen?", vbYesNo + vbQuestion, Programmanaam) = vbNo) Then
            Screen.MousePointer = vbDefault
            Exit Sub
        End If

        'Update: Notes()
        'Verwijder alle notes in de groep en subgroepen.
        For iNote = 1 To cntNotes
            If Left(Notes(iNote).Groep, Len(GroepDelete)) = GroepDelete Then
                Notes(iNote).ListNodeRemove
                Set Notes(iNote) = Nothing
            End If
        Next

        'Update: Notes()
        NotesUpdateArray
    End If

    'Update: lstItems
    frmMain.lstItems.Nodes.Remove (ConvertGroepToKey(GroepDelete))
    NodeChange frmMain.lstItems.SelectedItem

    'Update: AutoDeleteGroep
    AutoDeleteGroep = AutoDeleteGroepBackup

    'Update: IsDirty
    IsDirty = True

    'Update: MagInterfere; MousePointer
    Screen.MousePointer = vbDefault
    MagInterfere = True

    'Wilt u de notes nu opslaan?
    If AskSaveAfterDelete Then
        If MsgBox("Wilt u de notes nu opslaan?", vbYesNo + vbQuestion, Programmanaam) = vbYes Then NotesFileSave NotesFileOpen, NotesFileOpen & NotesFileBackupExtension
    End If

    Exit Sub
GroepDeleteError:
    ErrorHandling "modNotes", "GroepDelete", Err
    Resume Next
End Sub
'--------------------------------------------------------------------
'Deze procedure voegt een groep aan de lijst toe.
'--------------------------------------------------------------------
Public Sub GroepNew()
    On Error GoTo GroepNieuwError
    Dim NodeSelected As Node
    Dim NodeParentKey As String

    'Update: MagInterfere
    MagInterfere = False

    'Init: NodeSelected
    Set NodeSelected = frmMain.lstItems.SelectedItem

    'Check: NodeSelected
    Select Case Left(NodeSelected.Key, 3)
        Case KeyPrefixNote
            'Create: Groep
            NodeParentKey = NodeSelected.Parent.Key
            Set NodeSelected = frmMain.lstItems.Nodes.Add(NodeParentKey, tvwChild, NodeParentKey & "/Nieuwe Notegroep", "Nieuwe Notegroep", "groepDicht", "groepOpen")

        Case KeyPrefixGroep
            'Create: Groep
            NodeParentKey = NodeSelected.Key
            Set NodeSelected = frmMain.lstItems.Nodes.Add(NodeParentKey, tvwChild, NodeParentKey & "/Nieuwe Notegroep", "Nieuwe Notegroep", "groepDicht", "groepOpen")

        Case Else
            'Create: Groep
            Set NodeSelected = frmMain.lstItems.Nodes.Add(KeyRoot, tvwChild, KeyPrefixGroep & "Nieuwe Notegroep", "Nieuwe Notegroep", "groepDicht", "groepOpen")
    End Select

    'Update: NodeSelected
    NodeSelected.Selected = True
    NodeSelected.Expanded = True
    NodeChange NodeSelected

    'Update: MagInterfere
    MagInterfere = True

    'Wilt u de notes nu opslaan?
    If AskSaveAfterNew Then
        If MsgBox("Wilt u de notes nu opslaan?", vbYesNo + vbQuestion, Programmanaam) = vbYes Then NotesFileSave NotesFileOpen, NotesFileOpen & NotesFileBackupExtension
    End If

    Exit Sub
GroepNieuwError:
    Select Case Err.Number
        Case 35602
            MsgBox "U moet de groep ""Nieuwe Notegroep"" hernoemen of verwijderen voordat u een nieuwe groep kunt maken.", vbExclamation
            Err.Clear
        Case Else
            ErrorHandling "modNotes", "GroepNew", Err
    End Select
    Resume Next
End Sub
'--------------------------------------------------------------------
'Deze procedure verwijdert een note.
'--------------------------------------------------------------------
Public Sub NoteDelete()
    On Error GoTo NoteDeleteError
    Dim iNoteDelete As Integer

    'Weet u het zeker?
    If (MsgBox("Weet u zeker dat u deze note wilt verwijderen?", vbYesNo + vbQuestion, Programmanaam) = vbNo) Then Exit Sub

    'Init: iNoteDelete
    iNoteDelete = curNote

    'Update: Notes()
    Notes(iNoteDelete).ListNodeRemove
    Set Notes(iNoteDelete) = Nothing
    If Not (cntNotes = 1) Then
        If Not (iNoteDelete = cntNotes) Then
            Set Notes(iNoteDelete) = Notes(cntNotes)
            Notes(iNoteDelete).NoteIndex = iNoteDelete
            Set Notes(cntNotes) = Nothing
        End If
        cntNotes = cntNotes - 1
        ReDim Preserve Notes(1 To cntNotes)
    Else
        cntNotes = 0
        ReDim Notes(1 To 1)
        Set frmMain.lstItems.SelectedItem = frmMain.lstItems.Nodes(KeyRoot)
    End If
    NodeChange frmMain.lstItems.SelectedItem

    'Update: IsDirty
    IsDirty = True

    'Wilt u de notes nu opslaan?
    If AskSaveAfterDelete Then
        If MsgBox("Wilt u de notes nu opslaan?", vbYesNo + vbQuestion, Programmanaam) = vbYes Then NotesFileSave NotesFileOpen, NotesFileOpen & NotesFileBackupExtension
    End If

    Exit Sub
NoteDeleteError:
    ErrorHandling "modNotes", "NoteDeleteError", Err, True
    Resume Next
End Sub
'--------------------------------------------------------------------
'Deze procedure maakt een nieuwe note aan.
'--------------------------------------------------------------------
Public Sub NoteNieuw()
    On Error GoTo NoteNieuwError
    Dim NodeSelected As Node

    'Init: NodeSelected
    Set NodeSelected = frmMain.lstItems.SelectedItem

    'Update: Notes()
    cntNotes = cntNotes + 1
    ReDim Preserve Notes(1 To cntNotes)
    Set Notes(cntNotes) = New clsNote
    Notes(cntNotes).NoteIndex = cntNotes
    Select Case Left(NodeSelected.Key, 3)
        Case KeyPrefixNote
            Notes(cntNotes).Groep = ConvertKeyToGroep(NodeSelected.Parent.Key)
        Case KeyPrefixGroep
            Notes(cntNotes).Groep = ConvertKeyToGroep(NodeSelected.Key)
    End Select
    Notes(cntNotes).ListNodeAdd

    'Update: NodeSelected
    Set NodeSelected = Notes(cntNotes).Node
    NodeSelected.EnsureVisible
    NodeSelected.Selected = True
    NodeChange NodeSelected

    'Update: txtTekst
    With frmMain.txtTekst
        .Font.Name = FontDefaultName
        .Font.Size = FontDefaultSize
        .Font.Bold = FontDefaultBold
        .Font.Italic = FontDefaultItalic
        .Font.Underline = FontDefaultUnderline
        .Font.Strikethrough = FontDefaultStrikeThru
        .SelColor = FontDefaultColor
    End With

    'Update: IsDirty
    IsDirty = True
    
    'Wilt u de notes nu opslaan?
    If AskSaveAfterNew Then
        If MsgBox("Wilt u de notes nu opslaan?", vbYesNo + vbQuestion, Programmanaam) = vbYes Then NotesFileSave NotesFileOpen, NotesFileOpen & NotesFileBackupExtension
    End If

    Exit Sub
NoteNieuwError:
    ErrorHandling "modNotes", "NoteNieuw", Err, True
    Resume Next
End Sub
'--------------------------------------------------------------------
'Deze procedure laadt Notes uit NotesFile in Notes().
'--------------------------------------------------------------------
Public Function NotesFileLoad(ByVal NotesFile As String, Optional ByVal ClearNotes As Boolean = True) As Boolean
    On Error GoTo NotesFileLoadError
    Dim NotesFileVersion As Integer
    Dim EncryptionModifier As Long, Password As String, iLetterPassword As Integer
    Dim InputData As String
    Dim tmpTitel As String, tmpGroep As String, tmpZoekTermen As String, tmpZoekTermenArray() As String, tmpReserved1 As String, tmpReserved2 As String, tmpBeschrijving As String, tmpTekst As String
    Dim NodX As Node

    'Init: NotesFileLoad
    NotesFileLoad = True
    
    'Check: NotesFile
    If Dir(NotesFile) = "" Or NotesFile = "" Then
        MsgBox "Het notesbestand " & NotesFile & " bestaat niet. Selecteer een ander notesbestand om te openen.", vbExclamation
        Exit Function
    End If

    'Uw notes zijn gewijzigd. Wilt u de notes nu opslaan?
    If ClearNotes And IsDirty Then
        Select Case MsgBox("Uw notes zijn gewijzigd. Wilt u de notes nu opslaan?", vbYesNoCancel + vbQuestion)
            Case vbYes
                If NotesFileOpen = "" Then
                    If Not SelectFileSave(NotesFileOpen) Then Exit Function
                End If
                NotesFileSave NotesFileOpen, NotesFileOpen & NotesFileBackupExtension
            Case vbCancel
                Exit Function
        End Select
    End If

    'Init: Notes(); IsEncrypted
    If ClearNotes Then
        cntNotes = 0
        ReDim Notes(1 To 1)
        IsEncrypted = False
    End If

    'Open: NotesFile
    Open NotesFile For Input Access Read Lock Write As #1
        Do While Not EOF(1)
            'Init: InputData
            Line Input #1, InputData
            If Left(InputData, 3) = "// " Then InputData = ""
            InputData = Trim(InputData)

            'Check: InputData
            If InputData = "" Then GoTo NotesFileLoadVerder
            If NotesFileVersion = 0 Then
                If Left(InputData, 1) = "[" And Right(InputData, 1) = "]" Then
                    If Mid(InputData, 2, 2) = "NF" Then
                        If IsNumeric(Mid(InputData, 4, 1)) Then
                            'Init: NotesFileVersion
                            NotesFileVersion = CInt(Mid(InputData, 4, 1))
                        End If
                    End If
                End If
                If Not (NotesFileVersion = NotesFileCurrentVersion) Then _
                    Err.Raise vbObjectError + 1, , NotesFile & " is geen geldig notesbestand."
                GoTo NotesFileLoadVerder
            End If

            If InputData = "[ENCRYPTED]" Then
                'Show: frmPassword
                frmPassword.Show 1

                'Init: Password; EncryptionModifier
                Password = frmPassword.txtPassword
                EncryptionModifier = 0
                For iLetterPassword = 1 To Len(Password)
                    EncryptionModifier = EncryptionModifier + Asc(Mid(Password, iLetterPassword, 1))
                Next
                EncryptionModifier = ((EncryptionModifier Xor 26) Mod 11) + 2

                'Init: IsEncrypted
                If ClearNotes Then IsEncrypted = True

            ElseIf InputData = "<NOTE>" And Not EOF(1) Then
                'Init: tmpTitel, tmpGroep, tmpZoekTermen, tmpReserved1, tmpReserved2, tmpBeschrijving, tmpTekst
'                Input #1, tmpTitel, tmpGroep, tmpZoekTermen, tmpReserved1, tmpReserved2, tmpBeschrijving, tmpTekst
                
                Line Input #1, tmpTitel
                Line Input #1, tmpGroep
                Line Input #1, tmpZoekTermen
                Line Input #1, tmpReserved1
                Line Input #1, tmpReserved2
                Line Input #1, tmpBeschrijving
                Line Input #1, tmpTekst

                tmpTitel = ReplaceVariables(tmpTitel, True)
                tmpGroep = ReplaceVariables(tmpGroep, True)
                tmpZoekTermen = ReplaceVariables(tmpZoekTermen, True)
                'tmpReserved1 = ReplaceVariables(tmpReserved1, True)
                'tmpReserved2 = ReplaceVariables(tmpReserved2, True)
                tmpBeschrijving = ReplaceVariables(tmpBeschrijving, True)
                tmpTekst = ReplaceVariables(tmpTekst, True)
                If EncryptionModifier Then
                    tmpTitel = Decrypt(tmpTitel, EncryptionModifier)
                    tmpGroep = Decrypt(tmpGroep, EncryptionModifier)
                    tmpZoekTermen = Decrypt(tmpZoekTermen, EncryptionModifier)
                    'tmpReserved1 = Decrypt(tmpReserved1, EncryptionModifier)
                    'tmpReserved2 = Decrypt(tmpReserved2, EncryptionModifier)
                    tmpBeschrijving = Decrypt(tmpBeschrijving, EncryptionModifier)
                    tmpTekst = Decrypt(tmpTekst, EncryptionModifier)
                End If
                tmpZoekTermenArray = Split(tmpZoekTermen, ",")

                'Update: Notes()
                cntNotes = cntNotes + 1
                ReDim Preserve Notes(1 To cntNotes)
                Set Notes(cntNotes) = New clsNote
                With Notes(cntNotes)
                 .NoteIndex = cntNotes
                 .Titel = tmpTitel
                 .Groep = tmpGroep
                 .ZoekTermen = tmpZoekTermenArray
                 '.Reserved1 = tmpReserved1
                 '.Reserved2 = tmpReserved2
                 .Beschrijving = tmpBeschrijving
                 .Tekst = tmpTekst
                End With
            End If

NotesFileLoadVerder:
        Loop

    'Sluit het bestand.
    Close #1

    'Init: NotesFileOpen
    If ClearNotes Then NotesFileOpen = NotesFile

    'Update: mnuFile; tbrToolbar
    frmMain.mnuFileReload.Enabled = True
    frmMain.tbrToolbar.Buttons.Item("FileReload").Enabled = True

    'Init: lstItems
    If ClearNotes Then
        frmMain.lstItems.Nodes.Clear
        Set NodX = frmMain.lstItems.Nodes.Add(, , KeyRoot, Programmanaam, "INT_Desktop", "INT_Desktop")
        If Not (NodX Is Nothing) Then
            NodX.Expanded = ExpandRootNode
            NodX.Sorted = True
        End If
    End If

    'Update: lstItems
    For Each Note In Notes()
        If Note Is Nothing Then Exit For
        If Not Note.NodeInList Then Note.ListNodeAdd
    Next
    If frmMain.lstItems.SelectedItem Is Nothing Then
        Set frmMain.lstItems.SelectedItem = frmMain.lstItems.Nodes(KeyRoot)
        NodeChange frmMain.lstItems.SelectedItem
    End If

    'Update: Controls
    If ClearNotes Then
        frmMain.chkIsEncrypted.Value = IIf(IsEncrypted, vbChecked, vbUnchecked)
        frmMain.txtPassword.Enabled = IsEncrypted
        frmMain.txtPassword.Text = Password
    End If

    'Update: IsDirty
    IsDirty = Not ClearNotes

    Exit Function
NotesFileLoadError:
    Select Case Err.Number
        Case 62 'Input past end of file
            Err.Clear
            GoTo NotesFileLoadVerder
        Case 70
            MsgBox "Het notesbestand is op dit moment niet toegankelijk en waarschijnlijk in gebruik door een ander programma. De notes kunnen daardoor niet gelezen worden." & vbCrLf & "U kunt alle andere programma's afsluiten en het dan opnieuw proberen.", vbCritical + vbOKOnly
            Err.Clear
            NotesFileLoad = False
        Case vbObjectError + 1
            MsgBox Err.Description, vbCritical
            Err.Clear
            NotesFileLoad = False
        Case Else
            ErrorHandling "modNotes", "NotesFileLoad", Err, True
            NotesFileLoad = False
            Resume Next
    End Select
End Function
'--------------------------------------------------------------------
'Deze procedure maakt een nieuw, leeg notesbestand.
'--------------------------------------------------------------------
Public Sub NotesFileNew()
    On Error GoTo NotesFileNewError
    Dim NodX As Node

    'Uw notes zijn gewijzigd. Wilt u de notes nu opslaan?
    If IsDirty Then
        Select Case MsgBox("Uw notes zijn gewijzigd. Wilt u de notes nu opslaan?", vbYesNoCancel + vbQuestion)
            Case vbYes
                NotesFileSave NotesFileOpen, NotesFileOpen & NotesFileBackupExtension
            Case vbCancel
                Exit Sub
        End Select
    End If

    'Init: Notes(); IsEncrypted
    cntNotes = 0
    ReDim Notes(1 To 1)
    IsEncrypted = False

    'Init: NotesFileOpen
    NotesFileOpen = ""

    'Init: mnuFile; tbrToolbar
    frmMain.mnuFileReload.Enabled = False
    frmMain.tbrToolbar.Buttons.Item("FileReload").Enabled = False

    'Update: lstItems
    frmMain.lstItems.Nodes.Clear
    Set NodX = frmMain.lstItems.Nodes.Add(, , KeyRoot, Programmanaam, "INT_Desktop", "INT_Desktop")
    NodX.Expanded = ExpandRootNode
    NodX.Sorted = True
    NodX.Selected = True
    NodeChange NodX

    'Update: Controls
    frmMain.chkIsEncrypted.Value = IIf(IsEncrypted, vbChecked, vbUnchecked)
    frmMain.txtPassword.Enabled = IsEncrypted
    frmMain.txtPassword.Text = Password

    Exit Sub
NotesFileNewError:
    ErrorHandling "modNotes", "NotesFileNew", Err, True
    Resume Next
End Sub
'--------------------------------------------------------------------
'Deze procedure slaat de notes in het notesbestand op.
'--------------------------------------------------------------------
Public Sub NotesFileSave(ByVal NotesFile As String, Optional ByVal NotesFileBackup As String)
    On Error GoTo NotesFileSaveError
    Dim EncryptionModifier As Long
    Dim iNote As Integer

    'Copy: NotesFile -> NotesFileBackup
    If Not (NotesFileBackup = "") Then
        If Not (Dir(NotesFile) = "") Then
            FileCopy NotesFile, NotesFileBackup
        End If
    End If

    'Init: EncryptionModifier
    If IsEncrypted Then
        EncryptionModifier = 0
        For iLetterPassword = 1 To Len(Password)
            EncryptionModifier = EncryptionModifier + Asc(Mid(Password, iLetterPassword, 1))
        Next
        EncryptionModifier = ((EncryptionModifier Xor 26) Mod 11) + 2
    End If

    'Open: NotesFile
    Open NotesFile For Output Access Write Lock Write As #2
        'Write: Info
        Print #2, "// Dit is een notesbestand van " & Programmanaam & " Versie " & Versie & "."
        Print #2, "// " & Copyright
        Print #2,
        Print #2, "[NF" & NotesFileCurrentVersion & "]"
        If IsEncrypted Then Print #2, "[ENCRYPTED]"
        Print #2,

        'Write: Notes()
        If IsEncrypted Then
            For Each Note In Notes()
                If Note Is Nothing Then Exit For
                Print #2, "<NOTE>"

                Print #2, ReplaceVariables(Encrypt(Note.Titel, EncryptionModifier))
                Print #2, ReplaceVariables(Encrypt(Note.Groep, EncryptionModifier))
                Print #2, ReplaceVariables(Encrypt(Join(Note.ZoekTermen, ","), EncryptionModifier))
                Print #2, "Reserved"
                Print #2, "Reserved"
                Print #2, ReplaceVariables(Encrypt(Note.Beschrijving, EncryptionModifier))
                Print #2, ReplaceVariables(Encrypt(Note.Tekst, EncryptionModifier))

'                Write #2, ReplaceVariables(Encrypt(Note.Titel, EncryptionModifier)), _
                          ReplaceVariables(Encrypt(Note.Groep, EncryptionModifier)), _
                          ReplaceVariables(Encrypt(Join(Note.ZoekTermen, ","), EncryptionModifier)), _
                          "Reserved", _
                          "Reserved", _
                          ReplaceVariables(Encrypt(Note.Beschrijving, EncryptionModifier)), _
                          ReplaceVariables(Encrypt(Note.Tekst, EncryptionModifier))
                Print #2,
            Next
        Else
            For Each Note In Notes()
                If Note Is Nothing Then Exit For
                Print #2, "<NOTE>"

                Print #2, ReplaceVariables(Note.Titel)
                Print #2, ReplaceVariables(Note.Groep)
                Print #2, ReplaceVariables(Join(Note.ZoekTermen, ","))
                Print #2, "Reserved"
                Print #2, "Reserved"
                Print #2, ReplaceVariables(Note.Beschrijving)
                Print #2, ReplaceVariables(Note.Tekst)

'                Write #2, ReplaceVariables(Note.Titel), _
                          ReplaceVariables(Note.Groep), _
                          ReplaceVariables(Note.ZoekTermen), _
                          "Reserved", _
                          "Reserved", _
                          ReplaceVariables(Note.Beschrijving), _
                          ReplaceVariables(Note.Tekst)
                Print #2,
            Next
        End If

        'Write: Info
        Print #2, "// Dit notesbestand is gemaakt op " & Date & " @ " & Time & "."
        If cntNotes = 1 Then
            Print #2, "// Er is één note opgeslagen."
        Else
            Print #2, "// Er zijn " & cntNotes & " notes opgeslagen."
        End If

    'Close: NotesFile
    Close #2

    'Update: NotesFileOpen
    NotesFileOpen = NotesFile

    'Update: mnuFile; tbrToolbar
    frmMain.mnuFileReload.Enabled = True
    frmMain.tbrToolbar.Buttons.Item("FileReload").Enabled = True

    'Update: IsDirty
    IsDirty = False

    Exit Sub
NotesFileSaveError:
    Select Case Err.Number
        Case 53
            Err.Clear
        Case Else
            ErrorHandling "modNotes", "NotesFileSave", Err, True
    End Select
    Resume Next
End Sub
'--------------------------------------------------------------------
'Deze procedure maakt een nieuw notesbestand als die er nog niet is.
'--------------------------------------------------------------------
Public Sub NotesFileSaveNew()
    On Error GoTo NotesFileSaveNewError
    Dim FirstNoteBeschrijving As String, FirstNoteTekst As String

    'Init: FirstNoteBeschrijving; FirstNoteTekst
    FirstNoteBeschrijving = "Welkom bij " & Programmanaam & " !"
    FirstNoteTekst = "{\rtf1\ansi\deff0\deftab720{\fonttbl{\f0\fswiss MS Sans Serif;}{\f1\froman\fcharset2 Symbol;}{\f2\fswiss Arial;}{\f3\fswiss Tahoma;}}" & vbCrLf & _
                     "{\colortbl\red0\green0\blue0;}" & vbCrLf & _
                     "\deflang1033\pard\qc\plain\f3\fs44\cf0\b " & Programmanaam & "\plain\f3\fs17" & vbCrLf & _
                     "\par \pard\plain\f3\fs17" & vbCrLf & _
                     "\par \plain\f3\fs17\b " & Programmanaam & "\plain\f3\fs17  is een heel handig programma om notities en aantekeningen te bewaren. Je kan heel snel een nieuwe notitie maken door op \'e9\'e9n van de knoppen in de knoppenbalk te klikken en je kan een \plain\f3\fs17\b onbeperkt aantal notities\plain\f3\fs17  maken! Om alles overzichtelijk te houden worden de notities weergegeven in een handige boomstructuur waarin je groepen maken om de notities in te plaatsen. Ook kunt u door deze boomstructur makkelijk overschakelen tussen de verschillende notitites." & vbCrLf & _
                     "\par Je kan gebruik maken van alle bekende opmaakmethoden om de notities er zo goed mogelijk uit te laten zien. Door middel van een snelle \plain\f3\fs17\b zoekfunctie\plain\f3\fs17  is het mogelijk om elke notitie altijd heel snel terug te vinden." & vbCrLf & _
                     "\par" & vbCrLf & _
                     "\par Er kunnen meerdere \plain\f3\fs17\b notesbestanden\plain\f3\fs17  gebruikt worden, die elk met een \plain\f3\fs17\b wachtwoord\plain\f3\fs17  versleuteld kunnen worden." & vbCrLf & _
                     "\par" & vbCrLf & _
                     "\par Druk op \plain\f3\fs17\b F1\plain\f3\fs17  om meer hulp te krijgen met \plain\f3\fs17\b " & Programmanaam & "\plain\f3\fs17 ." & vbCrLf & _
                     "\par" & vbCrLf & _
                     "\par" & vbCrLf & _
                     "\par Als je deze notitie hebt gelezen dan kan je hem verwijderen door op \plain\f3\fs17\b Shift+Del\plain\f3\fs17  te drukken." & vbCrLf & _
                     "\par" & vbCrLf & _
                     "\par }"

    'Open: NotesFileName
    Open Directory & NotesFileNameDefault & NotesFileExtension For Output Access Write Lock Write As #3

        'Write: Info
        Print #3, "// Dit is een notesbestand van " & Programmanaam & " Versie " & Versie
        Print #3, "// " & Copyright
        Print #3,
        Print #3, "[NF" & NotesFileCurrentVersion & "]"
        Print #3,

        'Write: Note
        Print #3, "<NOTE>"

        Print #3, "Welkom bij " & Programmanaam & " !"
        Print #3, "Algemeen"
        Print #3, "Welkom"
        Print #3, "Reserved"
        Print #3, "Reserved"
        Print #3, ReplaceVariables(FirstNoteBeschrijving)
        Print #3, ReplaceVariables(FirstNoteTekst)

'        Write #3, "Welkom bij " & Programmanaam & " !", _
                  "Algemeen", _
                  "Welkom", _
                  "Reserved", _
                  "Reserved", _
                  ReplaceVariables(FirstNoteBeschrijving), _
                  ReplaceVariables(FirstNoteTekst)

        'Write: Info
        Print #3,
        Print #3, "// Dit bestand is gemaakt op " & Date & " @ " & Time & "."
        Print #3, "// Er is in dit bestand één note opgeslagen."

    'Close: NotesFileName
    Close #3

    Exit Sub
NotesFileSaveNewError:
    ErrorHandling "modNotes", "NotesFileSaveNew", Err
    Resume Next
End Sub
'--------------------------------------------------------------------
'Deze procedure werkt Notes() bij, zodat er geen 'gaten' van niet-
'gebruikte in blijven zitten.
'--------------------------------------------------------------------
Public Sub NotesUpdateArray()
    On Error GoTo NotesUpdateArrayError
    Dim iNote As Integer
    Dim iNoteLast As Integer

    'Update: Notes()
    For iNote = 1 To cntNotes
        If Notes(iNote) Is Nothing Then
            If Not (cntNotes = 1) And Not (iNote = cntNotes) Then
                For iNoteLast = iNote + 1 To cntNotes
                    If Not (Notes(iNoteLast) Is Nothing) Then
                        'Update: Notes()
                        Set Notes(iNote) = Notes(iNoteLast)
                        Notes(iNote).NoteIndex = iNote
                        Set Notes(iNoteLast) = Nothing

                        'Update: iNote
                        iNote = iNote - 1

                        Exit For
                    End If
                Next
            End If
        End If
    Next
    For iNote = 1 To cntNotes
        If Notes(iNote) Is Nothing Then
            'Update: Notes() Bound
            If Not (cntNotes = 1) Then
                cntNotes = iNote - 1
                ReDim Preserve Notes(1 To cntNotes)
            Else
                cntNotes = 0
                ReDim Notes(1 To 1)
            End If
            Exit For
        End If
    Next

    Exit Sub
NotesUpdateArrayError:
    ErrorHandling "modNotes", "NotesUpdateArray", Err, True
    Resume Next
End Sub
'--------------------------------------------------------------------
'Deze procedure wijzigt de naam van een groep.
'--------------------------------------------------------------------
Public Sub GroepWijzigNaam(ByVal GroepNaamNieuw As String)
    On Error GoTo GroepWijzigNaamError
    Dim NodeSelected As Node
    Dim GroepOud As String, GroepNieuw As String
    Dim AutoDeleteGroepBackup As Boolean
    Dim iNote As Integer, NoteGroep As String

    'Init: NodeSelected
    Set NodeSelected = frmMain.lstItems.SelectedItem

    'Init: GroepOud; GroepNieuw
    GroepNaamOud = ConvertKeyToGroep(NodeSelected.Key)
    GroepNieuw = Left(GroepNaamOud, InStrRev(GroepNaamOud, "\"))
    GroepNieuw = GroepNieuw & GroepNaamNieuw

    If GroepHasNotes(GroepNaamOud) = True Then
        'Update: AutoDeleteGroep
        AutoDeleteGroepBackup = AutoDeleteGroep
        AutoDeleteGroep = True

        'Update: Notes()
        For iNote = 1 To cntNotes
            NoteGroep = Notes(iNote).Groep
            If Left(NoteGroep, Len(GroepNaamOud)) = GroepNaamOud Then
                NoteGroep = Mid(NoteGroep, Len(GroepNaamOud) + 1)
                NoteGroep = GroepNieuw & NoteGroep
                Notes(iNote).Groep = NoteGroep
            End If
        Next

        'Update: AutoDeleteGroep
        AutoDeleteGroep = AutoDeleteGroepBackup
    Else
        'Update: NodeSelected
        NodeSelected.Key = ConvertGroepToKey(GroepNieuw)
        NodeSelected.Text = GroepNaamNieuw
        NodeSelected.Parent.Sorted = True
    End If

    'Update: IsDirty
    IsDirty = True

    Exit Sub
GroepWijzigNaamError:
    ErrorHandling "modNotes", "GroepWijzigNaam", Err
    Resume Next
End Sub
'--------------------------------------------------------------------
'Deze procedure werkt Notes() bij als er in een tekstvak iets wordt veranderd.
'--------------------------------------------------------------------
Public Sub WijzigNote(ByVal Property As String, ByVal Value As String)
    On Error GoTo WijzigNoteError

    'Check: MagInterfere; curNote
    If MagInterfere = False Then Exit Sub
    If curNote <= 0 Then Exit Sub

    'Update: Notes()
    Select Case Property
        Case "Titel"
            Notes(curNote).Titel = Value
            IsDirty = True
        Case "Groep"
            Notes(curNote).Groep = Value
            IsDirty = True
        Case "Beschrijving"
            Notes(curNote).Beschrijving = Value
            IsDirty = True
        Case "Tekst"
            Notes(curNote).Tekst = Value
            IsDirty = True
    End Select

    Exit Sub
WijzigNoteError:
    Select Case Err.Number
        Case 91
            Err.Clear
            Exit Sub
        Case Else
            ErrorHandling "modNotes", "WijzigNote", Err, True
            Resume Next
    End Select
End Sub
'--------------------------------------------------------------------
'Deze procedure laadt de note-ShowInfo zien, en laat de gebruiker de zoektermen wijzigen.
'--------------------------------------------------------------------
Public Sub LaadNoteInfo(ByVal ShowInfo As Boolean)
    On Error GoTo WijzigZoekTermenError
    Dim Termen() As String, cntTermen As Integer
    Dim iTerm As Integer

    'Check: curNote
    If Not (curNote > 0) Then Exit Sub

    'Init: Termen
    Termen = Notes(curNote).ZoekTermen
    cntTermen = Notes(curNote).ZoekTermenCount

    'Init: frmNoteInfo
    Load frmNoteInfo
    frmNoteInfo.txtNoteNaam.Text = Notes(curNote).Titel
    frmNoteInfo.txtNoteNummer.Text = curNote & " van " & cntNotes
    frmNoteInfo.txtNoteGroep.Text = Notes(curNote).Groep
    If cntTermen = 0 Then
        frmNoteInfo.txtNoteTrefwoorden.Text = "[geen]"
    Else
        frmNoteInfo.txtNoteTrefwoorden.Text = Join(Termen, ", ")
    End If
    For iTerm = 0 To cntTermen - 1
        frmNoteInfo.lstTermen.AddItem Termen(iTerm)
    Next

    Select Case ShowInfo
        Case True
            Set frmNoteInfo.tabTabStrip.SelectedItem = frmNoteInfo.tabTabStrip.Tabs.Item(1)
        Case False
            Set frmNoteInfo.tabTabStrip.SelectedItem = frmNoteInfo.tabTabStrip.Tabs.Item(2)
    End Select

    'Show: frmNoteInfo
    frmNoteInfo.Show 1
    If frmNoteInfo.Canceled Then
        Unload frmNoteInfo
        Exit Sub
    End If

    'Update: Notes()
    cntTermen = frmNoteInfo.lstTermen.ListCount
    ReDim Termen(0 To IIf(cntTermen = 0, 1, cntTermen) - 1)
    For iTerm = 0 To cntTermen - 1
        Termen(iTerm) = frmNoteInfo.lstTermen.List(iTerm)
    Next
    If Not (Join(Termen) = Join(Notes(curNote).ZoekTermen)) Then
        IsDirty = True
        Notes(curNote).ZoekTermen = Termen
    End If

    'Unload: frmNoteInfo
    Unload frmNoteInfo

    Exit Sub
WijzigZoekTermenError:
    ErrorHandling "modNotes", "WijzigZoekTermen", Err
    Resume Next
End Sub
