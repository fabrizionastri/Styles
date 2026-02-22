Option Explicit

' ============================
' Entry point
' ============================
Public Sub RemapStyleCleanup()
    Dim doc As Document
    Set doc = ActiveDocument
    
    Application.ScreenUpdating = False
    
    ' 1) Import styles from Normal.dotm
    ImportStylesFromNormalDotm doc
    
    ' 2) Remap styles per STYLE_ID_MAPPING
    RemapStyles doc
    
    ' 3) Delete unused styles
    DeleteUnusedStyles_Stronger doc
    
    Application.ScreenUpdating = True
    
    MsgBox "Done: imported Normal.dotm styles, remapped styles, and deleted unused styles.", vbInformation
End Sub

' ============================
' 1) Import from Normal.dotm
' ============================
Private Sub ImportStylesFromNormalDotm(ByVal doc As Document)
    Dim normalPath As String
    normalPath = Application.NormalTemplate.FullName
    
    ' Copies ALL styles (and related formatting building blocks) from Normal.dotm into the active doc.
    ' Note: This will overwrite existing styles in the document if names collide.
    doc.CopyStylesFromTemplate normalPath
End Sub

' ============================
' 2) Remap styles (your mapping)
' ============================
Private Sub RemapStyles(ByVal doc As Document)
    ' Mapping from OLD -> NEW
    ' STYLE_ID_MAPPING = {
    '   'Title': 'Heading1',
    '   'Heading 1': 'Article 1',
    '   'Heading 2': 'Article 2',
    '   'Heading 3': 'Article 3',
    '   'Heading 4': 'Article 4',
    '   'Heading 6': 'Heading 4',
    '   'PublishedOn': 'Published',
    '   'ToDo': 'Offset',
    ' }
    
    Dim m As Object
    Set m = CreateObject("Scripting.Dictionary")
    m.CompareMode = vbTextCompare ' case-insensitive mapping
    
    m.Add "Title", "Heading 1"
    m.Add "Heading 1", "Article 1"
    m.Add "Heading 2", "Article 2"
    m.Add "Heading 3", "Article 3"
    m.Add "Heading 4", "Article 4"
    m.Add "Heading 6", "Heading 4"
    m.Add "PublishedOn", "Published"
    m.Add "ToDo", "Offset"
    
    ' Ensure destination styles exist (import should provide them; otherwise we fail fast with a clear message)
    Dim k As Variant
    For Each k In m.keys
        EnsureStyleExists doc, CStr(m(k))
    Next k
    
    ' Apply remap across:
    ' - Paragraphs (most important for your Heading/Article mapping)
    ' - Character styles (for in-line text)
    ' - Tables, headers/footers, footnotes/endnotes, textboxes/shapes (StoryRanges)
    RemapAcrossAllStories doc, m
    
    ' Optional: also update ListTemplates / numbering is not directly remapped here (style change preserves numbering if tied to styles)
End Sub

Private Sub EnsureStyleExists(ByVal doc As Document, ByVal styleName As String)
    On Error GoTo ErrHandler
    Dim s As Style
    Set s = doc.Styles(styleName)
    Exit Sub
ErrHandler:
    Err.Clear
    MsgBox "Missing destination style: '" & styleName & "'. " & vbCrLf & _
           "Make sure it exists in Normal.dotm or in the document before running.", vbCritical
    Err.Raise vbObjectError + 513, "EnsureStyleExists", "Missing style: " & styleName
End Sub

Private Sub RemapAcrossAllStories(ByVal doc As Document, ByVal mapping As Object)
    Dim story As Range
    For Each story In doc.StoryRanges
        RemapInRange story, mapping
        
        ' Some stories have linked ranges (e.g., multiple headers)
        Do While Not story.NextStoryRange Is Nothing
            Set story = story.NextStoryRange
            RemapInRange story, mapping
        Loop
    Next story
End Sub

Private Sub RemapInRange(ByVal rng As Range, ByVal mapping As Object)
    Dim p As Paragraph
    Dim oldName As String, newName As String
    
    ' Paragraph styles
    For Each p In rng.Paragraphs
        oldName = StyleNameSafe(p.Style)
        If mapping.Exists(oldName) Then
            newName = CStr(mapping(oldName))
            On Error Resume Next
            p.Style = ActiveDocument.Styles(newName)
            On Error GoTo 0
        End If
    Next p
    
    ' Character styles (in-line). We scan by Words to avoid iterating every character.
    ' If you have heavy documents and this is too slow, we can optimize with Find per-style.
    Dim w As Range
    For Each w In rng.Words
        oldName = StyleNameSafe(w.Style)
        If mapping.Exists(oldName) Then
            newName = CStr(mapping(oldName))
            On Error Resume Next
            w.Style = ActiveDocument.Styles(newName)
            On Error GoTo 0
        End If
    Next w
End Sub

Private Function StyleNameSafe(ByVal styleVariant As Variant) As String
    ' styleVariant is often a Style object; sometimes it can be a string
    On Error Resume Next
    StyleNameSafe = ""
    If IsObject(styleVariant) Then
        StyleNameSafe = styleVariant.NameLocal
    Else
        StyleNameSafe = CStr(styleVariant)
    End If
    On Error GoTo 0
End Function

' ============================
' 3) Delete unused styles (stronger)
' ============================
Private Sub DeleteUnusedStyles_Stronger(ByVal doc As Document)
    ' This version counts usage across:
    ' - Paragraphs and in-line text (Words)
    ' - All StoryRanges (including headers/footers/footnotes/textboxes)
    ' It deletes unused non-built-in styles (paragraph + character + linked styles),
    ' but skips special cases and any style currently in use.
    
    Dim used As Object
    Set used = CreateObject("Scripting.Dictionary")
    used.CompareMode = vbTextCompare
    
    CollectUsedStylesAcrossAllStories doc, used
    
    Dim i As Long
    Dim s As Style
    Dim nm As String
    
    ' Delete from end to start to avoid index shift
    For i = doc.Styles.count To 1 Step -1
        Set s = doc.Styles(i)
        nm = s.NameLocal
        
        If CanDeleteStyle(s) Then
            If Not used.Exists(nm) Then
                On Error Resume Next
                s.Delete
                On Error GoTo 0
            End If
        End If
    Next i
End Sub

Private Sub CollectUsedStylesAcrossAllStories(ByVal doc As Document, ByVal used As Object)
    Dim story As Range
    Dim p As Paragraph
    Dim w As Range
    Dim nm As String
    
    For Each story In doc.StoryRanges
        ' Paragraph styles
        For Each p In story.Paragraphs
            nm = StyleNameSafe(p.Style)
            If Len(nm) > 0 Then used(nm) = True
        Next p
        
        ' Character styles
        For Each w In story.Words
            nm = StyleNameSafe(w.Style)
            If Len(nm) > 0 Then used(nm) = True
        Next w
        
        ' linked ranges
        Do While Not story.NextStoryRange Is Nothing
            Set story = story.NextStoryRange
            
            For Each p In story.Paragraphs
                nm = StyleNameSafe(p.Style)
                If Len(nm) > 0 Then used(nm) = True
            Next p
            
            For Each w In story.Words
                nm = StyleNameSafe(w.Style)
                If Len(nm) > 0 Then used(nm) = True
            Next w
        Loop
    Next story
End Sub

Private Function CanDeleteStyle(ByVal s As Style) As Boolean
    ' Conservative deletion rules.
    ' - Never delete built-in styles
    ' - Only delete paragraph/character/linked
    ' - Avoid deleting "Normal" even if not detected (it often is)
    ' - Avoid deleting styles that Word treats as special / throws errors on delete
    
    On Error GoTo SafeFalse
    
    If s.BuiltIn Then
        CanDeleteStyle = False
        Exit Function
    End If
    
    If s.Type <> wdStyleTypeParagraph And s.Type <> wdStyleTypeCharacter And s.Type <> wdStyleTypeLinked Then
        CanDeleteStyle = False
        Exit Function
    End If
    
    If StrComp(s.NameLocal, "Normal", vbTextCompare) = 0 Then
        CanDeleteStyle = False
        Exit Function
    End If
    
    CanDeleteStyle = True
    Exit Function
    
SafeFalse:
    CanDeleteStyle = False
End Function

