Attribute VB_Name = "modMain"
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
'Global Constants
'--------------------------------------------------------------------
Global Const Programmanaam As String = "WMS Notes"
Global Const Company As String = "Maaskant Software (WMS)"
Global Const Copyright As String = "Copyright © 1998 - 2004 Wout Maaskant"
Global Const Versie As String = "3.02.157"

'--------------------------------------------------------------------
'API Stuff
'--------------------------------------------------------------------
Public Declare Function GetTickCount Lib "Kernel32.dll" () As Long
Public Declare Function ShellExecute Lib "Shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function SHFileExists Lib "Shell32.dll" Alias "#45" (ByVal szPath As String) As Long

'Dit is voor de hyperlinks
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long

Public Type POINTAPI
   X As Long
   Y As Long
End Type

'--------------------------------------------------------------------
'Variabelen
'--------------------------------------------------------------------
Public Directory As String
Public Const InDutch As Boolean = True
Public MagInterfere As Boolean

Public ExpandRootNode As Boolean
Public AutoDeleteGroep As Boolean
Public AskSaveAfterDelete As Boolean
Public AskSaveAfterNew As Boolean

Public RegistrySavePosition As Boolean
'--------------------------------------------------------------------
'Deze procedure leegt en reset de tekstvelden.
'--------------------------------------------------------------------
Public Sub ClearFields()
    On Error GoTo ClearFieldsError

    'Update: frmMain
    frmMain.Caption = Programmanaam

    'Update: txtTitel; txtBeschrijving; txtTekst
    frmMain.txtTitel.Text = ""
    frmMain.txtTitel.ToolTipText = ""
    frmMain.txtBeschrijving.Text = ""
    frmMain.txtTekst.TextRTF = ""

    'Update: txtTekst
    With frmMain.txtTekst
        .SelStart = 0
        .SelLength = Len(.Text)

        'Default font zeg maar
        .BulletIndent = 284
        .SelAlignment = 0
        .SelBullet = False
        .SelProtected = False

        .SelFontName = FontDefaultName
        .SelFontSize = FontDefaultSize
        .SelBold = FontDefaultBold
        .SelItalic = FontDefaultItalic
        .SelUnderline = FontDefaultUnderline
        .SelStrikeThru = FontDefaultStrikeThru
        .SelColor = FontDefaultColor

        .SelCharOffset = 0
        .SelHangingIndent = 0
        .SelIndent = 0
        .SelRightIndent = 0
        .SelTabCount = 0
        .RightMargin = 0
    End With

    Exit Sub
ClearFieldsError:
    ErrorHandling "modMain", "ClearFields", Err
    Resume Next
End Sub
'--------------------------------------------------------------------
'Deze procedure converteert een NoteGroep naar een NodeKey.
'--------------------------------------------------------------------
Public Function ConvertGroepToKey(ByVal NoteGroep As String) As String
    On Error GoTo ConvertGroepToKeyError
    Dim NodeKey As String

    'Init: NodeKey
    NodeKey = Replace(NoteGroep, "\", "/")
    NodeKey = KeyPrefixGroep & NodeKey

    'Init: ConvertGroepToKey
    ConvertGroepToKey = NodeKey

    Exit Function
ConvertGroepToKeyError:
    ErrorHandling "modMain", "ConvertGroepToKey", Err, True
    Resume Next
End Function
'--------------------------------------------------------------------
'Deze procedure converteert een NodeKey naar een NoteGroep.
'--------------------------------------------------------------------
Public Function ConvertKeyToGroep(ByVal NodeKey As String) As String
    On Error GoTo ConvertKeyToGroepError
    Dim NoteGroep As String

    'Init: NoteGroep
    NoteGroep = Replace(NodeKey, "/", "\")
    If Left$(NodeKey, 3) = KeyPrefixGroep Then
        NoteGroep = Mid(NoteGroep, 4)
    End If

    'Init: ConvertKeyToGroep
    ConvertKeyToGroep = NoteGroep

    Exit Function
ConvertKeyToGroepError:
    ErrorHandling "modMain", "ConvertKeyToGroep", Err, True
    Resume Next
End Function
'--------------------------------------------------------
'Deze procedure versleutelt Tekst met Modifier.
'--------------------------------------------------------
Public Function Decrypt(ByVal Tekst As String, ByVal Modifier As Long) As String
    On Error GoTo DecryptError
    Dim iLetter As Long
    Dim ascLetter As Byte

    'Decrypt: Tekst
    For iLetter = 1 To Len(Tekst)
        ascLetter = Asc(Mid(Tekst, iLetter, 1))
        If ascLetter > 25 And ascLetter < 240 Then
            ascLetter = ascLetter Xor (iLetter Mod Modifier)
            If ascLetter > 25 And ascLetter < 240 Then
                Mid(Tekst, iLetter, 1) = Chr(ascLetter)
            End If
        End If
    Next iLetter

    'Update: Decrypt
    Decrypt = Tekst

    Exit Function
DecryptError:
    ErrorHandling "modMain", "Decrypt", Err
    Resume Next
End Function
'--------------------------------------------------------
'Deze procedure versleutelt Tekst met Modifier.
'--------------------------------------------------------
Public Function Encrypt(ByVal Tekst As String, ByVal Modifier As Long) As String
    On Error GoTo EncryptError
    Dim iLetter As Long
    Dim ascLetter As Byte

    'Encrypt: Tekst
    For iLetter = 1 To Len(Tekst)
        ascLetter = Asc(Mid(Tekst, iLetter, 1))
        If ascLetter > 25 And ascLetter < 240 Then
            ascLetter = ascLetter Xor (iLetter Mod Modifier)
            If ascLetter > 25 And ascLetter < 240 Then
                Mid(Tekst, iLetter, 1) = Chr(ascLetter)
            End If
        End If
    Next iLetter

    'Update: Encrypt
    Encrypt = Tekst

    Exit Function
EncryptError:
    ErrorHandling "modMain", "Encrypt", Err
    Resume Next
End Function
'--------------------------------------------------------------------------------
'Dit is het Enhanced Error System II (EES2) versie 3.1.
'--------------------------------------------------------------------------------
Public Sub ErrorHandling(ByVal Module As String, ByVal Functie As String, ByVal Fout As ErrObject, Optional ByVal Kritiek As Boolean = False, Optional ByVal InfoTekst As String = "")
    Dim Tekst As String

    If InDutch Then
        'Init: Tekst
        Tekst = "Module: " & Module & vbCrLf & _
                 "Functie: " & Functie & vbCrLf & _
                 "Error: " & Fout.Number & " - " & Fout.Description & _
                 vbCrLf & _
                 vbCrLf & _
                 IIf(InfoTekst <> "", InfoTekst, "Er is een fout in " & IIf(Kritiek = True, "een kritiek deel van ", "") & Programmanaam & " opgetreden.") & _
                 vbCrLf & _
                 vbCrLf & _
                 "Neem alstublieft zo spoedig mogelijk contact op met " & Company & " door te e-mailen naar " & Mail & " of door naar " & Website & " te gaan." & _
                 vbCrLf & _
                 vbCrLf & _
                 "Het is mogelijk dat u normaal door kunt gaan, maar " & Programmanaam & " kan ook onvoorspelbaar gaan reageren." & _
                 vbCrLf & _
                 "Wilt u doorgaan ?"
    Else
        'Init: Tekst
        Tekst = "Module: " & Module & vbCrLf & _
                 "Function: " & Functie & vbCrLf & _
                 "Error: " & Fout.Number & " - " & Fout.Description & _
                 vbCrLf & _
                 vbCrLf & _
                 IIf(InfoTekst <> "", InfoTekst, "An error occurred in " & IIf(Kritiek = True, "a critical part of ", "") & Programmanaam & ".") & _
                 vbCrLf & _
                 vbCrLf & _
                 "Please contact " & Company & " as soon as possible by e-mailing to " & Mail & " or by visiting " & WebsiteEN & "." & _
                 vbCrLf & _
                 vbCrLf & _
                 "It is possible that you can continue normally, but " & Programmanaam & " may give unpredictable reactions." & _
                 vbCrLf & _
                 "Would you like to continue?"
    End If

    'Laat het foutbericht zien.
    If MsgBox(Tekst, vbCritical + vbYesNo, Programmanaam & " - " & IIf(InDutch, "Fout", "Error")) = vbYes Then _
        Fout.Clear Else End
End Sub
'--------------------------------------------------------------------
'Deze procedure initialiseert imlTreeView, zodat die door lstItems
'gebruikt kan worden. Als er externe afbeeldingen aanwezig zijn, dan
'worden die geladen, en anders worden de interne afbeeldingen gebruikt.
'--------------------------------------------------------------------
Public Sub ImageListInit()
    On Error GoTo ImageListInitError

    'Groep: dicht.
    frmMain.imlTreeView.ListImages.Remove "groepDicht"
    If Dir(Directory & "GDicht.bmp") = "" Then
        Set ImgX = frmMain.imlTreeView.ListImages.Add(, "groepDicht", frmMain.imlTreeView.ListImages.Item("INT_GDicht").Picture)
    Else
        Set ImgX = frmMain.imlTreeView.ListImages.Add(, "groepDicht", LoadPicture(Directory & "GDicht.bmp"))
    End If

    'Groep: open.
    frmMain.imlTreeView.ListImages.Remove "groepOpen"
    If Dir(Directory & "GOpen.bmp") = "" Then
        Set ImgX = frmMain.imlTreeView.ListImages.Add(, "groepOpen", frmMain.imlTreeView.ListImages.Item("INT_GOpen").Picture)
    Else
        Set ImgX = frmMain.imlTreeView.ListImages.Add(, "groepOpen", LoadPicture(Directory & "GOpen.bmp"))
    End If

    'Node: dicht.
    frmMain.imlTreeView.ListImages.Remove "fileClosed"
    If Dir(Directory & "FDicht.bmp") = "" Then
        Set ImgX = frmMain.imlTreeView.ListImages.Add(, "fileClosed", frmMain.imlTreeView.ListImages.Item("INT_NDicht").Picture)
    Else
        Set ImgX = frmMain.imlTreeView.ListImages.Add(, "fileClosed", LoadPicture(Directory & "NDicht.bmp"))
    End If

    'Node: open.
    frmMain.imlTreeView.ListImages.Remove "fileOpen"
    If Dir(Directory & "FOpen.bmp") = "" Then
        Set ImgX = frmMain.imlTreeView.ListImages.Add(, "fileOpen", frmMain.imlTreeView.ListImages.Item("INT_NOpen").Picture)
    Else
        Set ImgX = frmMain.imlTreeView.ListImages.Add(, "fileOpen", LoadPicture(Directory & "NOpen.bmp"))
    End If

    Exit Sub
ImageListInitError:
    Select Case Err.Number
        Case 35601
            Err.Clear
        Case Else
            ErrorHandling "modMain", "ImageListInit", Err
    End Select
    Resume Next
End Sub
'--------------------------------------------------------------------
'Deze procedure start het programma op.
'--------------------------------------------------------------------
Public Sub Main()
    On Error GoTo MainError
    Dim StartTickCount As Long
    Dim CommandString As String, CommandImport As Boolean, CommandNotesFile As String

    'Update: MousePointer
    Screen.MousePointer = vbHourglass

    'Init: frmIntro
    frmIntro.Show
    frmIntro.Refresh
    DoEvents

    'Init: Introscherm Timer
    StartTickCount = GetTickCount

    'Init: Variabelen; HTMLHelp
    Directory = App.Path & IIf(Right(App.Path, 1) <> "\", "\", "")
'    If GetThreadLocale = LocaleDutch Then InDutch = True Else InDutch = False
    App.HelpFile = Directory & App.EXEName & IIf(InDutch, "_nl", "_en") & ".chm"
    HTMLHelp.DefaultHelpFile = App.HelpFile
    HTMLHelp.DefaulthWnd = frmMain.hWnd
    HTMLHelp.DefaultWindow = "Main"

    'Init: frmMain
    Load frmMain
    frmMain.Caption = Programmanaam

    'Initialiseer imlTreeView voor lstItems.
    ImageListInit

    'Init: Setup; WindowPosition
    RegistrySetupLoad
    RegistryPositionLoad

    'Init: CommandImport; CommandNotesFile
    CommandString = Trim(Command())
    If LCase(Left(CommandString, 2)) = "-i" Then
        CommandImport = True
        CommandNotesFile = LTrim(Mid(CommandString, 3))
    Else
        CommandImport = False
        CommandNotesFile = CommandString
    End If

    'Load: Notes
    If (CommandNotesFile = "") Or (Dir(CommandNotesFile) = "") Then
        NotesFileLoad NotesFileDefault
    Else
        If CommandImport = False Then
            NotesFileLoad CommandNotesFile
        Else
            NotesFileLoad NotesFileDefault
            NotesFileLoad CommandNotesFile, False

            'Update: IsDirty
            IsDirty = True
        End If
    End If

    ''Zorg dat er minimaal drie seconden gewacht wordt.
    'Do While GetTickCount - StartTickCount < 3000
    '    DoEvents
    'Loop

    'Update: frmMain; frmIntro
    frmMain.Show
    Unload frmIntro

    'Update: MousePointer; MagInterfere
    MagInterfere = True
    Screen.MousePointer = vbDefault

    Exit Sub
MainError:
    Select Case Err.Number
        Case 339, 713
            Err.Clear
            MsgBox Programmanaam & " heeft een aantal ActiveX controls nodig die niet (juist) zijn geïnstallerd op deze computer." & vbCrLf & _
                   "Download deze controls op " & Website & "programmas/files/vb6extra.exe en installeer ze." & vbCrLf & _
                   vbCrLf & _
                   Programmanaam & " zal worden afgesloten.", vbCritical
                   End
        Case 35602
            Err.Clear
        Case Else
            ErrorHandling "modMain", "Main", Err, True
    End Select
    Resume Next
End Sub
'--------------------------------------------------------------------
'Deze procedure sluit het programma af.
'
'Als MainEnd True geeft: Annuleren.
'--------------------------------------------------------------------
Public Function MainEnd() As Boolean
    On Error GoTo MainEndError

    If IsDirty Then
        'Uw notes zijn gewijzigd. Wilt u de notes nu opslaan?
        Select Case MsgBox("Uw notes zijn gewijzigd. Wilt u de notes opslaan voordat " & Programmanaam & " afgesloten wordt?", vbYesNoCancel + vbQuestion)
            Case vbYes
                NotesFileSave NotesFileOpen, NotesFileOpen & NotesFileBackupExtension
            Case vbCancel
                MainEnd = True
                Exit Function
        End Select
    End If

    'Save: WindowPosition
    If RegistrySavePosition = True Then RegistryPositionSave

    'Update: WindowState
    frmMain.WindowState = vbMinimized

    End
    Exit Function
MainEndError:
    ErrorHandling "modMain", "MainEnd", Err
    End
End Function
'--------------------------------------------------------------------
'Deze procedure wordt door lstItems aangeroepen als er van Node wordt veranderd.
'--------------------------------------------------------------------
Public Sub NodeChange(ByVal Node As Node)
    On Error GoTo NodeChangeError

    'Update: MagInterfere
    MagInterfere = False

    Select Case Left(Node.Key, 3)
        Case KeyPrefixNote
            'Init: curNote
            curNote = CInt(Mid(Node.Key, 4))

            'Update: Controls
            frmMain.txtBeschrijving.Locked = False
            frmMain.txtTekst.Locked = False
            LaadNote Node.Key

            'Update: mnuFile
            frmMain.mnuFileNoteDelete.Enabled = True
            frmMain.mnuNoteEditName.Enabled = True
            frmMain.mnuFileGroepDelete.Enabled = False
            frmMain.mnuFont.Enabled = True
            frmMain.mnuNoteInfo.Enabled = True
            frmMain.mnuNoteEditProperties.Enabled = True

            'Update: tbrToolbar
            frmMain.tbrToolbar.Buttons.Item("FileNoteDelete").Enabled = True
            frmMain.tbrToolbar.Buttons.Item("FileGroepDelete").Enabled = False
            frmMain.tbrToolbar.Buttons.Item("NoteInfo").Enabled = True
            frmMain.tbrToolbar.Buttons.Item("NoteEditProperties").Enabled = True

        Case Else
            'Update: Controls
            frmMain.txtBeschrijving.Locked = True
            frmMain.txtTekst.Locked = True
            ClearFields

            'Update: mnuBestand
            frmMain.mnuFileNoteDelete.Enabled = False
            frmMain.mnuFont.Enabled = False
            frmMain.mnuNoteInfo.Enabled = False
            frmMain.mnuNoteEditProperties.Enabled = False

            'Update: tbrToolbar
            frmMain.tbrToolbar.Buttons.Item("FileNoteDelete").Enabled = False
            frmMain.tbrToolbar.Buttons.Item("NoteInfo").Enabled = False
            frmMain.tbrToolbar.Buttons.Item("NoteEditProperties").Enabled = False

            If Left(Node.Key, 3) = KeyPrefixGroep Then
                'Init: curNote
                curNote = 0

                'Update: frmMain
                frmMain.txtTitel.Text = Node.Text

                'Update: Controls
                frmMain.mnuNoteEditName.Enabled = True
                frmMain.mnuFileGroepDelete.Enabled = True
                frmMain.tbrToolbar.Buttons.Item("FileGroepDelete").Enabled = True

            ElseIf Node.Key = KeyRoot Then
                'Init: curNote
                curNote = -1

                'Update: frmMain
                frmMain.txtTitel.Text = ""

                'Update: Controls
                frmMain.mnuNoteEditName.Enabled = False
                frmMain.mnuFileGroepDelete.Enabled = False
                frmMain.tbrToolbar.Buttons.Item("FileGroepDelete").Enabled = False
            End If
    End Select

    'Update: MagInterfere
    MagInterfere = True

    Exit Sub
NodeChangeError:
    ErrorHandling "modMain", "NodeChange", Err, True
    Resume Next
End Sub
'--------------------------------------------------------------------
'Deze procedure laadt de vorige positie en afmeting van frmMain uit het register.
'--------------------------------------------------------------------
Public Sub RegistryPositionLoad()
    On Error GoTo RegistryPositionLoadError
    Dim Resultaat As Variant

    'Init: RegisterEdit
    RegisterEdit.KeyMain = CURRENT_USER
    RegisterEdit.KeySub = "Software\[deep]software\" & Company & "\" & Programmanaam

    'Read: WindowState
    Resultaat = RegisterEdit.ValueRead("WindowState", RegString)
    If Not IsEmpty(Resultaat) Then
        frmMain.WindowState = CInt(Resultaat)
    Else
        frmMain.WindowState = vbNormal
        frmMain.Left = CSng((Screen.Width - frmMain.Width) / 2)
        frmMain.Top = CSng((Screen.Height - frmMain.Height) / 2)
    End If

    'Check: WindowState
    ' Het werkt om altijd alle gegevens te laden en het geeft zelfs een
    ' beter effect. De regel frmMain.WindowState = x heeft pas effect
    ' als frmMain.Show wordt uitgevoerd. Als je de Left/Top/Width/Height
    ' probeert te wijzigen als Not(frmMain.WindowState = vbNormal), dan
    ' komt er een fout, maar nu dus niet omdat frmMain.WindowState = vbNormal.
    'If Not (frmMain.WindowState = vbNormal) Then Exit Sub

    'Read: WindowLeft
    Resultaat = RegisterEdit.ValueRead("WindowLeft", RegString)
    If Not IsEmpty(Resultaat) Then frmMain.Left = CSng(Resultaat)

    'Read: WindowTop
    Resultaat = RegisterEdit.ValueRead("WindowTop", RegString)
    If Not IsEmpty(Resultaat) Then frmMain.Top = CSng(Resultaat)

    'Read: WindowWidth
    Resultaat = RegisterEdit.ValueRead("WindowWidth", RegString)
    If Not IsEmpty(Resultaat) Then frmMain.Width = CSng(Resultaat)

    'Read: WindowHeight
    Resultaat = RegisterEdit.ValueRead("WindowHeight", RegString)
    If Not IsEmpty(Resultaat) Then frmMain.Height = CSng(Resultaat)

    'Read: SplitterLeft
    Resultaat = RegisterEdit.ValueRead("SplitterLeft", RegString)
    If Not IsEmpty(Resultaat) Then frmMain.SizeControls (CInt(Resultaat))

    Exit Sub
RegistryPositionLoadError:
'    ErrorHandling "modMain", "RegistryPositionLoad", Err
    Err.Clear
    Resume Next
End Sub
'--------------------------------------------------------------------
'Deze procedure verwijdert alle instellingen (contextmenus niet!) uit het register.
'--------------------------------------------------------------------
Public Sub RegistryDelete()
    On Error GoTo RegistryDeleteError

    'Init: RegisterEdit
    RegisterEdit.KeyMain = CURRENT_USER
    RegisterEdit.KeySub = "Software\[deep]software\" & Company

    'Delete: Programmanaam
    RegisterEdit.KeyDelete Programmanaam

    'Show: Message
    MsgBox "Alle instellingen van " & Programmanaam & " zijn uit het register verwijderd.", vbInformation

    Exit Sub
RegistryDeleteError:
    ErrorHandling "modMain", "RegistryDelete", Err
    Resume Next
End Sub
'--------------------------------------------------------------------
'Deze procedure verwijdert de contextmenus.
'--------------------------------------------------------------------
Public Sub RegistryContextDelete()
    On Error GoTo RegistryContextDeleteError

    'Init: RegisterEdit
    RegisterEdit.KeyMain = CLASSES_ROOT
    RegisterEdit.KeySub = ""

    'Delete: CLASSES_ROOT\.nf3
    RegisterEdit.KeyDelete NotesFileExtension

    'Delete: CLASSES_ROOT\WMS NotesFile
    RegisterEdit.KeyDelete Programmanaam & "File"

    Exit Sub
RegistryContextDeleteError:
    ErrorHandling "modMain", "RegistryContextDelete", Err
End Sub
'--------------------------------------------------------------------
'Deze procedure maakt contextmenus in Windows voor notesbestanden.
'--------------------------------------------------------------------
Public Sub RegistryContextCreate()
    On Error GoTo RegistryContextCreateError

    'Init: RegisterEdit
    RegisterEdit.KeyMain = CLASSES_ROOT

    'Save: CLASSES_ROOT\.nf3
    RegisterEdit.KeySub = NotesFileExtension
    RegisterEdit.ValueSet "", RegString, Programmanaam & "File"

    'Save: CLASSES_ROOT\WMS NotesFile
    RegisterEdit.KeySub = Programmanaam & "File"
    RegisterEdit.ValueSet "", RegString, "Notesbestand"

    'Save: CLASSES_ROOT\WMS NotesFile\DefaultIcon
    RegisterEdit.KeySub = Programmanaam & "File\DefaultIcon"
    RegisterEdit.ValueSet "", RegString, Directory & App.EXEName & ".exe,0"

    'Save: CLASSES_ROOT\WMS NotesFile\shell\open\command
    RegisterEdit.KeySub = Programmanaam & "File\shell\open\command"
    RegisterEdit.ValueSet "", RegString, Directory & App.EXEName & ".exe %1"

    'Save: CLASSES_ROOT\WMS NotesFile\shell\import
    RegisterEdit.KeySub = Programmanaam & "File\shell\import"
    RegisterEdit.ValueSet "", RegString, "&Importeren"

    'Save: CLASSES_ROOT\WMS NotesFile\shell\import\command
    RegisterEdit.KeySub = Programmanaam & "File\shell\import\command"
    RegisterEdit.ValueSet "", RegString, Directory & App.EXEName & ".exe -i %1"

    Exit Sub
RegistryContextCreateError:
    ErrorHandling "modMain", "RegistryContextCreate", Err
End Sub
'--------------------------------------------------------------------
'Deze procedure slaat de positie en afmeting van frmMain op in het register.
'--------------------------------------------------------------------
Public Sub RegistryPositionSave()
    On Error GoTo RegistryPositionSaveError

    'Check: frmMain
    If frmMain.WindowState = vbMinimized Then Exit Sub

    'Init: RegisterEdit
    RegisterEdit.KeyMain = CURRENT_USER
    RegisterEdit.KeySub = "Software\[deep]software\" & Company & "\" & Programmanaam

    'Save: WindowState
    RegisterEdit.ValueSet "WindowState", RegString, CStr(frmMain.WindowState)

    'Save: WindowLeft; WindowTop; WindowWidth; WindowHeight
    If frmMain.WindowState = vbNormal Then
        RegisterEdit.ValueSet "WindowLeft", RegString, CStr(frmMain.Left)
        RegisterEdit.ValueSet "WindowTop", RegString, CStr(frmMain.Top)
        RegisterEdit.ValueSet "WindowWidth", RegString, CStr(frmMain.Width)
        RegisterEdit.ValueSet "WindowHeight", RegString, CStr(frmMain.Height)
    End If

    'Save: SplitterLeft
    RegisterEdit.ValueSet "SplitterLeft", RegString, CStr(frmMain.lblSplitter.Left)

    Exit Sub
RegistryPositionSaveError:
    ErrorHandling "modMain", "RegistryPositionSave", Err
End Sub
'--------------------------------------------------------------------
'Deze procedure laadt de instellingen uit het register. Als die er
'niet zijn, dan worden standaard-waarden genomen. Deze procedure maakt
'eventueel ook een nieuw notesbestand.
'--------------------------------------------------------------------
Public Sub RegistrySetupLoad()
    On Error GoTo RegistrySetupLoadError
    Dim Resultaat As Variant
    Dim Temp As String

    'Init: RegisterEdit
    RegisterEdit.KeyMain = CURRENT_USER
    RegisterEdit.KeySub = "Software\[deep]software\" & Company & "\" & Programmanaam

    'Read: ExpandRootNode
    Resultaat = RegisterEdit.ValueRead("ExpandRootNode", RegString)
    If Not IsEmpty(Resultaat) Then _
        ExpandRootNode = CBool(Resultaat) Else _
        ExpandRootNode = True

    'Read: AutoDeleteGroep
    Resultaat = RegisterEdit.ValueRead("AutoDeleteGroep", RegString)
    If Not IsEmpty(Resultaat) Then _
        AutoDeleteGroep = CBool(Resultaat) Else _
        AutoDeleteGroep = False

    'Read: AskSaveAfterDelete
    Resultaat = RegisterEdit.ValueRead("AskSaveAfterDelete", RegString)
    If Not IsEmpty(Resultaat) Then _
        AskSaveAfterDelete = CBool(Resultaat) Else _
        AskSaveAfterDelete = False

    'Read: AskSaveAfterNew
    Resultaat = RegisterEdit.ValueRead("AskSaveAfterNew", RegString)
    If Not IsEmpty(Resultaat) Then _
        AskSaveAfterNew = CBool(Resultaat) Else _
        AskSaveAfterNew = False

    'Read: FontDefaultName
    Resultaat = RegisterEdit.ValueRead("FontDefaultName", RegString)
    If Not IsEmpty(Resultaat) Then
        If Not CStr(Resultaat) = "" Then _
            FontDefaultName = CStr(Resultaat) Else _
            FontDefaultName = "Arial"
    Else
        FontDefaultName = "Arial"
    End If

    'Read: FontDefaultSize
    Resultaat = RegisterEdit.ValueRead("FontDefaultSize", RegString)
    If Not IsEmpty(Resultaat) Then _
        FontDefaultSize = CSng(Resultaat) Else _
        FontDefaultSize = 8

    'Read: FontDefaultBold
    Resultaat = RegisterEdit.ValueRead("FontDefaultBold", RegString)
    If Not IsEmpty(Resultaat) Then _
        FontDefaultBold = CBool(Resultaat) Else _
        FontDefaultBold = False

    'Read: FontDefaultItalic
    Resultaat = RegisterEdit.ValueRead("FontDefaultItalic", RegString)
    If Not IsEmpty(Resultaat) Then _
        FontDefaultItalic = CBool(Resultaat) Else _
        FontDefaultItalic = False

    'Read: FontDefaultUnderline As Boolean
    Resultaat = RegisterEdit.ValueRead("FontDefaultUnderline", RegString)
    If Not IsEmpty(Resultaat) Then _
        FontDefaultUnderline = CBool(Resultaat) Else _
        FontDefaultUnderline = False

    'Read: FontDefaultStrikeThru
    Resultaat = RegisterEdit.ValueRead("FontDefaultStrikeThru", RegString)
    If Not IsEmpty(Resultaat) Then _
        FontDefaultStrikeThru = CBool(Resultaat) Else _
        FontDefaultStrikeThru = False

    'Read: FontDefaultColor
    Resultaat = RegisterEdit.ValueRead("FontDefaultColor", RegString)
    If Not IsEmpty(Resultaat) Then _
        FontDefaultColor = CLng(Resultaat) Else _
        FontDefaultColor = vbWindowText

    'Read: NotesFileDefault
    Resultaat = RegisterEdit.ValueRead("NotesFileDefault", RegString)
    If Not IsEmpty(Resultaat) Then NotesFileDefault = CStr(Resultaat)

    'Update: NotesFileDefault
    If Dir(NotesFileDefault) = "" Or NotesFileDefault = "" Then
        Temp = Dir(Directory & "*" & NotesFileExtension)
        If Not (Temp = "") Then
            NotesFileDefault = Directory & Temp
        Else
            NotesFileSaveNew
            NotesFileDefault = Directory & NotesFileNameDefault & NotesFileExtension
        End If
    End If

    'Read: RegistrySavePosition
    Resultaat = RegisterEdit.ValueRead("RegistrySavePosition", RegString)
    If Not IsEmpty(Resultaat) Then _
        RegistrySavePosition = CBool(Resultaat) Else _
        RegistrySavePosition = True

    Exit Sub
RegistrySetupLoadError:
    ErrorHandling "modMain", "RegistrySetupLoad", Err
    Resume Next
End Sub
'--------------------------------------------------------------------
'Deze procedure slaat de instellingen op in het register.
'--------------------------------------------------------------------
Public Sub RegistrySetupSave()
    On Error GoTo RegistrySetupSaveError

    'Init: RegisterEdit
    RegisterEdit.KeyMain = CURRENT_USER
    RegisterEdit.KeySub = "Software\[deep]software\" & Company & "\" & Programmanaam

    'Save: ExpandRootNode; AutoDeleteGroep; AskSaveAfterDelete; AskSaveAfterNew
    RegisterEdit.ValueSet "ExpandRootNode", RegString, IIf(ExpandRootNode, "1", "0")
    RegisterEdit.ValueSet "AutoDeleteGroep", RegString, IIf(AutoDeleteGroep, "1", "0")
    RegisterEdit.ValueSet "AskSaveAfterDelete", RegString, IIf(AskSaveAfterDelete, "1", "0")
    RegisterEdit.ValueSet "AskSaveAfterNew", RegString, IIf(AskSaveAfterNew, "1", "0")

    'Save: FontDefault*
    RegisterEdit.ValueSet "FontDefaultName", RegString, FontDefaultName
    RegisterEdit.ValueSet "FontDefaultSize", RegString, CStr(FontDefaultSize)
    RegisterEdit.ValueSet "FontDefaultBold", RegString, IIf(FontDefaultBold, "1", "0")
    RegisterEdit.ValueSet "FontDefaultItalic", RegString, IIf(FontDefaultItalic, "1", "0")
    RegisterEdit.ValueSet "FontDefaultUnderline", RegString, IIf(FontDefaultUnderline, "1", "0")
    RegisterEdit.ValueSet "FontDefaultStrikeThru", RegString, IIf(FontDefaultStrikeThru, "1", "0")
    RegisterEdit.ValueSet "FontDefaultColor", RegString, CStr(FontDefaultColor)

    'Save: NotesFileDefault
    RegisterEdit.ValueSet "NotesFileDefault", RegString, NotesFileDefault

    'Save: RegistrySavePosition
    RegisterEdit.ValueSet "RegistrySavePosition", RegString, IIf(RegistrySavePosition, "1", "0")

    Exit Sub
RegistrySetupSaveError:
    ErrorHandling "modMain", "RegistrySetupSave", Err
    Resume Next
End Sub
'--------------------------------------------------------------------
'Deze procedure vervangt een aantal variabelen door tekstwaarden en
'andersom. De procedures NotesFileLoad en NotesFileSave gebruiken dit.
'--------------------------------------------------------------------
Public Function ReplaceVariables(ByVal Tekst As String, Optional ByVal Restore As Boolean = False) As String
    On Error GoTo ReplaceVariablesError

    'Update: Tekst
    If Not Restore Then
'        Tekst = Replace(Tekst, """", "$''$")
        Tekst = Replace(Tekst, vbCrLf, "$@$")
'        Tekst = Replace(Tekst, Chr(13), "$#$")
'        Tekst = Replace(Tekst, Chr(10), "$%$")
        Tekst = Replace(Tekst, Chr(26), "$&$")
    Else
'        Tekst = Replace(Tekst, "$''$", """")
        Tekst = Replace(Tekst, "$@$", vbCrLf)
'        Tekst = Replace(Tekst, "$#$", Chr(13))
'        Tekst = Replace(Tekst, "$%$", Chr(10))
        Tekst = Replace(Tekst, "$&$", Chr(26))
    End If

    'Init: ReplaceVariables
    ReplaceVariables = Tekst

    Exit Function
ReplaceVariablesError:
    ErrorHandling "modMain", "ReplaceVariables", Err
    Resume Next
End Function
'--------------------------------------------------------------------
'Deze procedure laat de gebruiker een notesbestand selecteren om te openen.
'--------------------------------------------------------------------
Public Function SelectFileOpen(ByRef File As String) As Boolean
    On Error GoTo SelectFileOpenError

    'Init: SelectFileOpen
    SelectFileOpen = True

    'Init: cdlFile
    With frmMain.cdlFile
     .CancelError = True
     .FileName = ""
     If .InitDir = "" And Not (NotesFileOpen = "") Then .InitDir = Left(NotesFileOpen, InStrRev(NotesFileOpen, "\") - 1)
     .Filter = "Notesbestanden|*" & NotesFileExtension & "|" & _
               "Notesbackupbestanden|*" & NotesFileBackupExtension & "|"
     .Flags = cdlOFNFileMustExist Or _
              cdlOFNHideReadOnly Or _
              cdlOFNLongNames Or _
              cdlOFNPathMustExist
     .ShowOpen
     .InitDir = CurDir
    End With

    'Init: File
    File = frmMain.cdlFile.FileName

    Exit Function
SelectFileOpenError:
    Select Case Err.Number
        Case 0, cdlCancel
            Err.Clear
            File = ""
            SelectFileOpen = False
        Case Else
            ErrorHandling "modMain", "SelectFileOpen", Err
            Resume Next
    End Select
End Function
'--------------------------------------------------------------------
'Deze procedure laat de gebruiker een bestandsnaam kiezen om het notesbestand onder op te slaan.
'--------------------------------------------------------------------
Public Function SelectFileSave(ByRef File As String) As Boolean
    On Error GoTo SelectFileSaveError

    'Init: SelectFileSave
    SelectFileSave = True

    'Init: cdlFile
    With frmMain.cdlFile
     .CancelError = True
     .FileName = ""
     .Filter = "Notesbestanden|*" & NotesFileExtension & "|"
     .Flags = cdlOFNFileMustExist Or _
              cdlOFNHideReadOnly Or _
              cdlOFNLongNames Or _
              cdlOFNOverwritePrompt
     .ShowSave
     .InitDir = CurDir
    End With

    'Init: File
    File = frmMain.cdlFile.FileName

    Exit Function
SelectFileSaveError:
    Select Case Err.Number
        Case 0, cdlCancel
            Err.Clear
            File = ""
            SelectFileSave = False
        Case Else
            ErrorHandling "modMain", "SelectFileSave", Err
            Resume Next
    End Select
End Function
'--------------------------------------------------------------------
'Deze procedure laat de gebruiker de instellingen wijzigen.
'--------------------------------------------------------------------
Public Sub SetupEdit()
    On Error GoTo SetupEditError

With frmSetup

    'Init: frmSetup
    Load frmSetup

    ' ExpandRootNode; AutoDeleteGroep; AskSaveAfterDelete; AskSaveAfterNew
    .chkExpandRootNode.Value = IIf(ExpandRootNode, vbChecked, vbUnchecked)
    .chkAutoDeleteGroep.Value = IIf(AutoDeleteGroep, vbChecked, vbUnchecked)
    .chkAskSaveAfterDelete.Value = IIf(AskSaveAfterDelete, vbChecked, vbUnchecked)
    .chkAskSaveAfterNew.Value = IIf(AskSaveAfterNew, vbChecked, vbUnchecked)

    ' FontDefault*
    .lblFontDefault.Font.Name = FontDefaultName
    .lblFontDefault.Font.Size = FontDefaultSize
    .lblFontDefault.Font.Bold = FontDefaultBold
    .lblFontDefault.Font.Italic = FontDefaultItalic
    .lblFontDefault.Font.Underline = FontDefaultUnderline
    .lblFontDefault.Font.Strikethrough = FontDefaultStrikeThru
    .lblFontDefault.ForeColor = FontDefaultColor
    .lblFontDefault.Caption = FontDefaultName & " " & CInt(FontDefaultSize) & " pt"

    ' NotesFileDefault
    .txtNotesFileDefault.Text = NotesFileDefault

    ' Contextmenus; RegistrySavePosition
    .chkRegistryContextCreate.Value = vbUnchecked
    .chkRegistryPositionSave.Value = IIf(RegistrySavePosition, vbChecked, vbUnchecked)

    'Shop: frmSetup
    frmSetup.Show 1
    If frmSetup.Canceled Then Exit Sub

    'Update: ExpandRootNode; AutoDeleteGroep; AskSaveAfterDelete; AskSaveAfterNew
    ExpandRootNode = IIf(.chkExpandRootNode.Value = vbChecked, True, False)
    AutoDeleteGroep = IIf(.chkAutoDeleteGroep.Value = vbChecked, True, False)
    AskSaveAfterDelete = IIf(.chkAskSaveAfterDelete.Value = vbChecked, True, False)
    AskSaveAfterNew = IIf(.chkAskSaveAfterNew.Value = vbChecked, True, False)

    'Update: FontDefault*
    FontDefaultName = .lblFontDefault.Font.Name
    FontDefaultSize = .lblFontDefault.Font.Size
    FontDefaultBold = .lblFontDefault.Font.Bold
    FontDefaultItalic = .lblFontDefault.Font.Italic
    FontDefaultUnderline = .lblFontDefault.Font.Underline
    FontDefaultStrikeThru = .lblFontDefault.Font.Strikethrough
    FontDefaultColor = .lblFontDefault.ForeColor

    'Update: NotesFileDefault
    NotesFileDefault = frmSetup.txtNotesFileDefault.Text

    'Registry: Contextmenus; RegistrySavePosition; Setup
    If .chkRegistryContextCreate.Value = vbChecked Then RegistryContextCreate
    RegistrySavePosition = IIf(.chkRegistryPositionSave.Value = vbChecked, True, False)
    If .chkRegistrySetupSave.Value = vbChecked Then RegistrySetupSave

End With

    Exit Sub
SetupEditError:
    ErrorHandling "modMain", "SetupEdit", Err
    Resume Next
End Sub
'--------------------------------------------------------------------
'Deze procedure zoekt in de trefwoorden en in de beschrijvingen van
'de notities naar een bepaald woord of een stuk tekst.
'--------------------------------------------------------------------
Public Function Zoeken(ByVal SearchText As String, ByVal SearchTrefwoorden As Boolean, ByVal SearchDescription As Boolean, ByVal IdentiekeLetters As Boolean, ByVal HeelWoord As Boolean) As Boolean
    On Error GoTo ZoekenError

    Dim Search() As String, cntFind As Integer 'Hier wordt op gezocht
    Dim Words() As String 'De woorden in de beschrijving
    Dim Termen() As String 'De zoektermen van de items

    Dim iNote As Integer

    Dim Match As Boolean 'Komen er termen overeen met het zoekwoord?
    Dim MatchTermen As String 'De termen die overeenkomen met de zoekwoorden

    Dim LVIX As ListItem

    'Init: Zoeken
    Zoeken = False

    'Update: SearchText
    SearchText = Replace(Trim(SearchText), ", ", ",")

    'Init: Search()
    Search() = Split(SearchText, ",")
    cntFind = UBound(Search) - LBound(Search) + 1

    'Check: Notes()
    For iNote = 1 To cntNotes
        MatchTermen = ""

        For Each Item In Search()
            'Compare: Termen() <-> Words()
            If SearchDescription Then
                'Init: Words()
                Words() = Split(Notes(iNote).Beschrijving, " ")

                For Each Word In Words
                    'Init: Match
                    Match = False
                    If IdentiekeLetters = True And HeelWoord = True Then
                        'Hoofd/kleine letters moeten Match zijn, én het moet 1 woord zijn.
                        If Item = Word Then Match = True

                    ElseIf IdentiekeLetters = False And HeelWoord = True Then
                        'Het moet 1 woord zijn.
                        If UCase$(Item) = UCase$(Word) Then Match = True

                    ElseIf IdentiekeLetters = True And HeelWoord = False Then
                        'Hoofd/kleine letters moeten Match zijn.
                        If Not (InStr(1, Word, Item, vbBinaryCompare) = 0) Then Match = True

                    ElseIf IdentiekeLetters = False And HeelWoord = False Then
                        'De hoofd/kleine letters hoeven niet te kloppen, en het hoeft niet één woord te zijn.
                        If Not (InStr(1, Word, Item, vbTextCompare) = 0) Then Match = True
                    End If

                    'Init: MatchTermen
                    If Match Then
                        If Not (MatchTermen = "") Then MatchTermen = MatchTermen & ", "
                        MatchTermen = MatchTermen & "[beschrijving]"
                        Exit For
                    End If
                Next Word
            End If

            'Compare: Termen() <-> Search()
            If SearchTrefwoorden Then
                'Init: Termen()
                Termen() = Notes(iNote).ZoekTermen

                For Each Term In Termen()
                    'Init: Match
                    Match = False
                    If IdentiekeLetters = True And HeelWoord = True Then
                        'Hoofd/kleine letters moeten Match zijn, én het moet 1 woord zijn.
                        If Item = Term Then Match = True

                    ElseIf IdentiekeLetters = False And HeelWoord = True Then
                        'Het moet 1 woord zijn.
                        If UCase$(Item) = UCase$(Term) Then Match = True

                    ElseIf IdentiekeLetters = True And HeelWoord = False Then
                        'Hoofd/kleine letters moeten Match zijn.
                        If Not (InStr(1, Term, Item, vbBinaryCompare) = 0) Then Match = True

                    ElseIf IdentiekeLetters = False And HeelWoord = False Then
                        'De hoofd/kleine letters hoeven niet te kloppen, en het hoeft niet één woord te zijn.
                        If Not (InStr(1, Term, Item, vbTextCompare) = 0) Then Match = True
                    End If

                    'Init: MatchTermen
                    If Match = True Then
                        If Not (MatchTermen = "") Then MatchTermen = MatchTermen & ", "
                        MatchTermen = MatchTermen & Term
                    End If
                Next Term
            End If
        Next Item

        If Not (MatchTermen = "") Then
            'Init: lvwResults
            Set LVIX = frmZoeken.lvwResults.ListItems.Add(, KeyPrefixNote & iNote, Notes(iNote).Titel)
            LVIX.SubItems(1) = MatchTermen
            Set LVIX = Nothing

            'Update: Zoeken
            Zoeken = True

            'Update: MatchTermen
            MatchTermen = ""
        End If

ZoekenVolgendeNote:
    Next

    'Init: lvwResults
    If Zoeken = False Then frmZoeken.lvwResults.ListItems.Add , "Nothing", "Geen trefwoorden gevonden"

    Exit Function
ZoekenError:
    ErrorHandling "modMain", "Zoeken", Err
    Resume Next
End Function
'--------------------------------------------------------------------
'Deze procedure scheidt een String naar een array van Strings op elk
'punt waar een komma (",") voorkomt. Deze procedure is vervangen door
'de VB functie Split().
'ZoekenMaakArray(var) = Split(var, ",") 'Volgens mij geeft Split een Variant met subtype String.
'--------------------------------------------------------------------
Public Function ZoekenMaakArray(ByVal ZoekTekst As String) As String()
'    On Error GoTo ZoekenMaakArrayError
'    Dim Items() As String
'    Dim cntItems As Integer
'
'    'Init: ZoekTekst; Items()
'    ZoekTekst = ZoekTekst & ","
'    cntItems = 0
'    ReDim Items(1 To 1) As String
'
'    'Init: Items()
'    Do
'        'Update: Items()
'        cntItems = cntItems + 1
'        ReDim Preserve Items(1 To cntItems) As String
'        Items(cntItems) = Left(ZoekTekst, InStr(1, ZoekTekst, ",") - 1)
'
'        'Update: ZoekTekst
'        ZoekTekst = Mid(ZoekTekst, InStr(1, ZoekTekst, ",") + 1)
'    Loop Until ZoekTekst = ""
'
'    'Init: ZoekenMaakArray
'    ZoekenMaakArray = Items()
'
'    Exit Function
'ZoekenMaakArrayError:
'    ErrorHandling "modMain", "ZoekenMaakArray", Err
'    Resume Next
End Function
