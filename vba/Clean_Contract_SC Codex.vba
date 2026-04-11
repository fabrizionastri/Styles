Sub Clean_Contract_SC()
    ' Normalizes spacing and applies color formatting/NoProofing to template tags:
    '   {{ ... }} and ‹ ... › = Blue + No Spellcheck → auto & manual data tags)
    '   {% ... %} and [ ... ] = Green + No Spellcheck → auto & manual logic tags)
    '   {? ... ?}   = Grey  → comments, to be removed before finalizing contract

    Dim rng As Range
    Set rng = ActiveDocument.Content

    ' ===== STEP 1: Normalize spacing for {{ }} and {% %} =====
    ' (Standardizing spacing first ensures the wildcards in Step 3 catch everything)
    
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchWildcards = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        
        ' Normalize {{
        .Text = "{{"
        .Replacement.Text = "{{ "
        .Execute Replace:=wdReplaceAll
        
        ' Normalize }}
        .Text = "}}"
        .Replacement.Text = " }}"
        .Execute Replace:=wdReplaceAll

        
        ' Normalize %}
        .Text = "%}"
        .Replacement.Text = " %}"
        .Execute Replace:=wdReplaceAll

        ' Normalize {%
        .Text = "{%"
        .Replacement.Text = "{% "
        .Execute Replace:=wdReplaceAll

        ' Remplace double spaces with single spaces (in case of multiple tags)
        .Text = "  "
        .Replacement.Text = " "
        .Execute Replace:=wdReplaceAll
        
    End With

    Set rng = ActiveDocument.Content

    ' Set document language before disabling proofing on template tags.
    ActiveDocument.Content.LanguageID = wdEnglishUK

    ' ===== STEP 2: Apply Colors and Disable Spellcheck (NoProofing) =====

    ' --- Grey: {? ... ?} (No proofing changes) ---
    Call ApplyDelimitedTagFormat(rng, "{?", "?}", RGB(128, 128, 128), False)

    ' --- Blue: {{ ... }} (No Spellcheck) ---
    Call ApplyDelimitedTagFormat(rng, "{{", "}}", wdColorBlue, True)

    ' --- Blue: ‹ ... › (No Spellcheck) ---
    Call ApplyDelimitedTagFormat(rng, "‹", "›", wdColorBlue, True)

    ' --- Green: {% ... %} (No Spellcheck) ---
    Call ApplyDelimitedTagFormat(rng, "{%", "%}", RGB(0, 176, 80), True)

    ' --- Green: [ ... ] (No Spellcheck) ---
    Call ApplyDelimitedTagFormat(rng, "[", "]", RGB(0, 176, 80), True)

    MsgBox "Cleaning Complete. Logic fields are now excluded from Spellcheck.", vbInformation
End Sub

' Helper Subroutine to keep the main code clean
Sub ApplyDelimitedTagFormat(ByRef rng As Range, openText As String, closeText As String, textColor As Long, disableProofing As Boolean)
    Dim searchRng As Range
    Dim closeRng As Range
    Dim tagRng As Range

    Set searchRng = rng.Duplicate

    Do
        With searchRng.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text = openText
            .Forward = True
            .Wrap = wdFindStop
            .Format = False
            .MatchWildcards = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With

        If Not searchRng.Find.Execute Then Exit Do

        Set closeRng = rng.Duplicate
        closeRng.SetRange searchRng.End, rng.End

        With closeRng.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text = closeText
            .Forward = True
            .Wrap = wdFindStop
            .Format = False
            .MatchWildcards = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With

        If Not closeRng.Find.Execute Then Exit Do

        Set tagRng = rng.Duplicate
        tagRng.SetRange searchRng.Start, closeRng.End
        tagRng.Font.Color = textColor
        tagRng.NoProofing = disableProofing

        searchRng.SetRange closeRng.End, rng.End
    Loop
End Sub
