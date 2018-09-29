VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
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
'Variabelen
'--------------------------------------------------------------------
Private mvarNoteIndex As Integer
Private mvarTitel As String
Private mvarGroep As String
Private mvarZoekTermen() As String
'Public ZoekTermen As String
Public Beschrijving As String
Public Tekst As String

Private mvarNodeInList As Boolean

Private mvarNode As Node
Private Sub Class_Initialize()
    mvarTitel = "Nieuwe Note"
    mvarGroep = "Nieuw"
    mvarZoekTermen = Split("", ",")
    Beschrijving = "Hier kunt u de beschrijving van de note invoeren."
    Tekst = "Hier kunt u uw note invoeren."
    mvarNodeInList = False
    mvarNoteIndex = 0
End Sub
Public Property Let Groep(ByVal vData As String)
    On Error GoTo GroepLetError
    Dim GroepOud As String, GroepNieuw As String

    'Init: GroepOud; GroepNieuw
    GroepOud = mvarGroep
    GroepNieuw = vData

    'Update: mvarGroep
    mvarGroep = GroepNieuw

    'Update: lstItems
    If mvarNodeInList = True Then
        'Create: GroepNieuw
        ListGroepCreate GroepNieuw

        'Update: mvarNode
        Set mvarNode.Parent = frmMain.lstItems.Nodes(ConvertGroepToKey(GroepNieuw))

        'Delete: GroepOud
        ListGroepDelete GroepOud
    End If

    Exit Property
GroepLetError:
    Select Case Err.Number
        Case 91, 35602
            Err.Clear
        Case Else
            ErrorHandling "Note", "Groep Let", Err
    End Select
    Resume Next
End Property
Public Property Get Groep() As String
    Groep = mvarGroep
End Property
Public Property Get Node() As Node
    Set Node = mvarNode
End Property
Public Property Get NodeInList() As Boolean
    NodeInList = mvarNodeInList
End Property
Public Property Let NoteIndex(ByVal vData As Integer)
    'Update: mvarNoteIndex
    mvarNoteIndex = vData

    'Update: mvarNode
    If mvarNodeInList Then
        mvarNode.Key = KeyPrefixNote & mvarNoteIndex
    End If
End Property
Public Property Get NoteIndex() As Integer
    NoteIndex = mvarNoteIndex
End Property
Public Property Let Titel(ByVal vData As String)
    On Error Resume Next

    'Init: mvarTitel
    mvarTitel = vData

    'Update: mvarNode
    If mvarNodeInList = True Then
        mvarNode.Text = mvarTitel
        mvarNode.Parent.Sorted = True
        frmMain.txtTitel.Text = mvarTitel
        frmMain.Caption = Programmanaam & " - " & mvarTitel
    End If
End Property
Public Property Get Titel() As String
    Titel = mvarTitel
End Property
Public Property Let ZoekTermen(ByRef vData() As String)
    'Update: mvarZoekTermen
    mvarZoekTermen = vData
End Property
Public Property Get ZoekTermen() As String()
    ZoekTermen = mvarZoekTermen
End Property
Public Property Get ZoekTermenCount() As Integer
    On Error Resume Next

    ZoekTermenCount = UBound(mvarZoekTermen) - LBound(mvarZoekTermen) + 1
End Property
Private Sub ListGroepCreate(ByVal NoteGroepCreate As String)
    On Error GoTo ListGroepCreateError
    Dim GroepCreate As String, GroepCreateParent As String
    Dim posSeparator As Integer
    Dim NodX As Node, NodeKey As String

    'Check: NoteGroepCreate
    Do Until NoteGroepCreate = ""
        'Init: GroepCreate // Update: NoteGroepCreate
        posSeparator = InStr(1, NoteGroepCreate, "\")
        If posSeparator = 0 Then posSeparator = Len(NoteGroepCreate) + 1
        GroepCreate = Left(NoteGroepCreate, posSeparator - 1)
        NoteGroepCreate = Mid(NoteGroepCreate, posSeparator + 1)

        'Create: Nodes
        If GroepCreateParent = "" Then
            'GroepCreate is een hoofdgroep.
            NodeKey = KeyPrefixGroep & GroepCreate
            Set NodX = frmMain.lstItems.Nodes.Add(KeyRoot, tvwChild, NodeKey, GroepCreate, "groepDicht", "groepOpen")
            GroepCreateParent = GroepCreate
        Else
            'GroepCreate is een subgroep.
            NodeKey = KeyPrefixGroep & GroepCreateParent & "/" & GroepCreate
            Set NodX = frmMain.lstItems.Nodes.Add(KeyPrefixGroep & GroepCreateParent, tvwChild, NodeKey, GroepCreate, "groepDicht", "groepOpen")
            GroepCreateParent = GroepCreateParent & "/" & GroepCreate
        End If
        NodX.Sorted = True
    Loop

    Exit Sub
ListGroepCreateError:
    Select Case Err.Number
        Case 91, 35602
            Err.Clear
        Case Else
            ErrorHandling "Note", "ListGroepCreate", Err
    End Select
    Resume Next
End Sub
Private Sub ListGroepDelete(ByVal NoteGroepDelete As String)
    On Error GoTo ListGroepDeleteError
    Dim NodeKeyDelete As String
    Dim posSeparator As Integer

    'Check: AutoDeleteGroep
    If Not AutoDeleteGroep Then Exit Sub

    'Init: NodeKeyDelete
    NodeKeyDelete = ConvertGroepToKey(NoteGroepDelete)

    If frmMain.lstItems.Nodes.Item(NodeKeyDelete).Children < 1 Then
        'Update: NodeKeyDelete
        'Kijk welke groep maximaal verwijderd mag worden.
        Do While frmMain.lstItems.Nodes.Item(NodeKeyDelete).Parent.Children <= 1
            posSeparator = InStrRev(NodeKeyParent, "/")
            If posSeparator = 0 Then Exit Do
            NodeKeyParent = Left(NodeKeyParent, posSeparator - 1)
        Loop

        'Delete: NodeKeyDelete
        frmMain.lstItems.Nodes.Remove NodeKeyDelete
    End If

    Exit Sub
ListGroepDeleteError:
    Select Case Err.Number
        Case 91, 35602
            Err.Clear
        Case Else
            ErrorHandling "Note", "ListGroepDelete", Err
    End Select
    Resume Next
End Sub
Public Sub ListNodeAdd()
    On Error GoTo ListNodeAddError
    Dim NodX As Node, NodeKey As String

    'Check: mvarNodeInList; mvarTitel; mvarGroep; mvarNoteIndex
    If mvarNodeInList Then Exit Sub
    If mvarTitel = "" Or mvarGroep = "" Or mvarNoteIndex = 0 Then Exit Sub

    'Create: Node(KeyRoot)
    Set NodX = frmMain.lstItems.Nodes.Add(, , KeyRoot, Programmanaam, "INT_Desktop", "INT_Desktop")
    If Not (NodX Is Nothing) Then
        NodX.Expanded = ExpandRootNode
        NodX.Sorted = True
    End If

    'Create: mvarGroep; mvarNode
    ListGroepCreate mvarGroep
    NodeKey = KeyPrefixNote & mvarNoteIndex
    Set mvarNode = frmMain.lstItems.Nodes.Add(ConvertGroepToKey(mvarGroep), tvwChild, NodeKey, mvarTitel, "fileClosed", "fileOpen")

    'Update: mvarNodeInList
    mvarNodeInList = True

    Exit Sub
ListNodeAddError:
    Select Case Err.Number
        Case 91, 35602
            Err.Clear
        Case Else
            ErrorHandling "Note", "ListNodeAdd", Err
    End Select
    Resume Next
End Sub
Public Sub ListNodeRemove()
    On Error GoTo ListNodeRemoveError
    Dim NodeKey As String
    Dim NodeKeyParent As String
    Dim posSeparator As Integer

    'Check: mvarNodeInList
    If Not mvarNodeInList Then Exit Sub

    'Delete: mvarNode
    frmMain.lstItems.Nodes.Remove mvarNode.Key
    Set mvarNode = Nothing
    ListGroepDelete mvarGroep

    'Update: mvarNodeInList
    mvarNodeInList = False

    Exit Sub
ListNodeRemoveError:
    Select Case Err.Number
        Case 91, 35602
            Err.Clear
        Case Else
            ErrorHandling "Note", "ListNodeRemove", Err
    End Select
    Resume Next
End Sub
