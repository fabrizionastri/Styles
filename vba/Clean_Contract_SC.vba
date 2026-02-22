Sub Clean_Contract_SC()
'
' Clean_Contract_SC Macro
' Normalizes spacing and applies color formatting to template tags:
'   {{ ... }} = Blue
'   {% ... %} = Green
'   { ... }   = Grey
'

    ' ===== STEP 1: Normalize spacing for {{ }} =====

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting

    ' Add space after {{
    With Selection.Find
        .Text = "{{"
        .Replacement.Text = "{{ "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    ' Replace [ and < by ‹
    With Selection.Find
        .Text = "["
        .Replacement.Text = "‹"
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    With Selection.Find
        .Text = "<"
        .Replacement.Text = "‹"
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    ' Replace ] and > by ›
    With Selection.Find
        .Text = "]"
        .Replacement.Text = "›"
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    With Selection.Find
        .Text = ">"
        .Replacement.Text = "›"
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    ' Remove double space after {{
    With Selection.Find
        .Text = "{{  "
        .Replacement.Text = "{{ "
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    ' Add space before }}
    With Selection.Find
        .Text = "}}"
        .Replacement.Text = " }}"
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    ' Remove double space before }}
    With Selection.Find
        .Text = "  }}"
        .Replacement.Text = " }}"
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    ' ===== STEP 2: Normalize spacing for {% %} =====

    ' Add space after {%
    With Selection.Find
        .Text = "{%"
        .Replacement.Text = "{% "
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    ' Remove double space after {%
    With Selection.Find
        .Text = "{%  "
        .Replacement.Text = "{% "
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    ' Add space before %}
    With Selection.Find
        .Text = "%}"
        .Replacement.Text = " %}"
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    ' Remove double space before %}
    With Selection.Find
        .Text = "  %}"
        .Replacement.Text = " %}"
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    ' ===== STEP 3: Apply colors (grey first, then blue/green overwrite) =====

    ' --- Grey: { ... } (single braces with spaces) ---
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Color = RGB(128, 128, 128)

    With Selection.Find
        .Text = "\{ ?* \}"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = True
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    ' --- Blue: {{ ... }} ---
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Color = wdColorBlue

    With Selection.Find
        .Text = "\{\{?*\}\}"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = True
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    ' --- Green: {% ... %} ---
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Color = RGB(0, 176, 80)

    With Selection.Find
        .Text = "\{\%?*\%\}"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = True
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    ' Reset find settings
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting


    '=== Apply British English language setting to entire document ===
    ActiveDocument.Content.LanguageID = wdEnglishUK


End Sub
