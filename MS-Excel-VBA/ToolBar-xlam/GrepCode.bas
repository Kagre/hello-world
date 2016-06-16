'Requires Reference: Microsoft VBScript Regular Expressions 5.5
Private Sub grep_tst()
    'MsgBox "]" & grep("xy", "[\~`!@#$%^&*()_\-+={}\[\]|\\;:<>,./?]") & "["
    MsgBox "]" & grep("xy", "[!#-&(-/:-@\[-`{-\~]") & "["
End Sub

Public Function grep(SearchText As String, PatternText As String, Optional MatchInstance As Long = 1) As String
    Dim re As New VBScript_RegExp_55.RegExp, m As VBScript_RegExp_55.Match, mCount As Long
    re.Global = MatchInstance <> 1 'set to true to find all occurances of pattern
    re.IgnoreCase = True
    re.MultiLine = False 'not sure; i assume it means look in multiple lines
                         '   for a match, but it could also mean that the match
                         '   may span multiple lines...
    re.Pattern = PatternText 'for examples see: http://regexlib.com/CheatSheet.aspx
    mCount = 1
    For Each m In re.Execute(SearchText)
        grep = m.Value 'the first match since global is false
        If mCount = MatchInstance Then Exit For
        mCount = mCount + 1
    Next m
End Function

Public Function countString(SearchText As String, PatternText As String) As Long
    Dim re As New VBScript_RegExp_55.RegExp, m As VBScript_RegExp_55.Match
    Dim mc As VBScript_RegExp_55.MatchCollection
    re.Global = True 'set to true to find all occurances of pattern
    re.IgnoreCase = True
    re.MultiLine = False 'not sure; i assume it means look in multiple lines
                         '   for a match, but it could also mean that the match
                         '   may span multiple lines...
    re.Pattern = PatternText 'for examples see: http://regexlib.com/CheatSheet.aspx
    Set mc = re.Execute(SearchText)
    countString = mc.count
'    For Each m In re.Execute(SearchText)
'        grep = m.Value 'the first match since global is false
'    Next m
End Function

Public Function REGEXP_LIKE(SearchText As String, PatternText As String) As Boolean
    REGEXP_LIKE = Len(grep(SearchText, PatternText)) > 0
End Function

Public Function FitsPattern(ToMatchString As String, Pattern As String) As Boolean
    FitsPattern = (grep(ToMatchString, Pattern) = ToMatchString)
End Function
