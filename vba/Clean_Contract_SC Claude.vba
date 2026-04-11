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

    ' ===== STEP 2: Apply Colors and Disable Spellcheck (NoProofing) =====

    ' --- Grey: { ... } (Single braces - No proofing changes) ---
    Call ApplyTagFormat(rng, "\{ ?* \}", RGB(128, 128, 128), False)

    ' --- Blue: {{ ... }} (No Spellcheck) ---
    Call ApplyTagFormat(rng, "\{\{?*\}\}", wdColorBlue, True)

    ' --- Blue: ‹ ... › (No Spellcheck) ---
    Call ApplyTagFormat(rng, "‹*›", wdColorBlue, True)

    ' --- Green: {% ... %} (No Spellcheck) ---
    Call ApplyTagFormat(rng, "\{\%?*\%\}", RGB(0, 176, 80), True)

    ' --- Green: [ ... ] (No Spellcheck) ---
    Call ApplyTagFormat(rng, "\[?*\]", RGB(0, 176, 80), True)

    ' Finalize: Set document language
    ActiveDocument.Content.LanguageID = wdEnglishUK
    
    MsgBox "Cleaning Complete. Logic fields are now excluded from Spellcheck.", vbInformation
End Sub

' Helper Subroutine to keep the main code clean
Sub ApplyTagFormat(ByRef rng As Range, findText As String, textColor As Long, disableProofing As Boolean)
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = findText
        .Replacement.Text = "" ' Keep found text
        .Replacement.Font.Color = textColor
        .Replacement.NoProofing = disableProofing
        
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
End Sub
