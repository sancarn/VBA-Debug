Public Type CSS_Style
    sName As String
    sStyle As String
End Type

Function cssToString(css As CSS_Style) As String
    'Css to string without vbCrLf
    'cssToString = "." & css.sName & " {" & css.sStyle & "}"
    
    Dim cssSubParts As String: cssSubParts = Replace(css.sStyle, ";", ";" & vbCrLf)
    cssSubParts = Left(cssSubParts, Len(cssSubParts) - Len(vbCrLf))
    cssToString = "." & css.sName & "{" & vbCrLf & _
        sTabs(cssSubParts, 1) & vbCrLf & _
    "}"
End Function

Sub removeCSS(css As CSS_Style, ByVal sName As String)
    'Takes some cssString e.g.
    '    "color:#FFFF00;background-color:#ff0000;float:left;"
    'and removes the value named sName. E.G. if sName = "background-color" then
    '    return "color:#FFFF00;float:left;"
    
    'Might want to loop this in case duplicates in css exist
    Dim iStt, iEnd As Integer
    iStt = InStr(1, ";" & css.sStyle, ";" & sName & ":") - 1
    iEnd = InStr(iStt + 1, css.sStyle, ";")
    css.sStyle = Left(css.sStyle, iStt) & Mid(css.sStyle, iEnd + 1)
End Sub

Sub addCSS(css As CSS_Style, ByVal sName As String, ByVal sValue As String)
    'sName - the name of the css 'type' that you want to add
    'sValue - the value of the css 'type' that you want to set
    'e.g:
    'background-color:#ff0000
    '    >>> sName = "background-color" , sValue = "#ff0000"
    
    css.sStyle = sName & ":" & sValue & ";" & css.sStyle
End Sub

Sub replaceCSS(css As CSS_Style, ByVal sName As String, sValue As String)
    'combines addCSS and removeCSS to replace existing CSS in a string
    
    removeCSS css, sName
    addCSS css, sName, sValue
End Sub

Sub replaceCSSColor(css As CSS_Style, ByVal sName As String, ByVal iColour As Integer)
    'Get hexadecimal value of rgb color
    Dim rgbHex As String
    rgbHex = Hex(iColour)
    rgbHex = String(6 - Len(rgbHex), "0") & rgbHex
    
    'Change hexadecimal string (including #)
    replaceCSS css, sName, "#" & rgbHex
End Sub

Function newCSSDefault(ByVal sName As String) As CSS_Style
    Dim css As CSS_Style
    css.sName = sName
    css.sStyle = "float:left;clear:both;"
    newCSSDefault = css
End Function

'Testing script
Sub test()
    'Initialise css
    Dim css As CSS_Style
    css = newCSSDefault("Error")
    
    'Add CSS
    addCSS css, "background-color", "#ff0000"
    addCSS css, "color", "#FFFF00"
    Debug.Print cssToString(css)
    
    'change html color
    replaceCSS css, "background-color", "red"
    Debug.Print cssToString(css)
    
    'add rgb color
    replaceCSSColor css, "background-color", RGB(255, 0, 0)
    Debug.Print cssToString(css)
    
End Sub

Function wrapInHTMLTag(ByVal toWrap As String, tag As String, Optional tagParams As String = "")
    'seperate tagParams from tag if exists
    If tagParams <> "" Then tagParams = " " & tagParams
    
    'return wrapped XML
    wrapInHTMLTag = "<" & tag & tagParams & ">" & vbCrLf & sTabs(toWrap, 1) & vbCrLf & "</" & tag & ">"
End Function

Function sTabs(ByVal toTab As String, ByVal numTabs As Integer) As String
    Dim tabStr As String: tabStr = String(numTabs, vbTab)
    sTabs = tabStr & Replace(toTab, vbCrLf, vbCrLf & tabStr)
End Function

Function sHTMLHeader(sTitle As String, sStyles As String, Optional miscTags As String = "")
    'create title and add (create?) styles
    Dim sInnerHTML As String
    sInnerHTML = wrapInHTMLTag(sTitle, "title") & vbCrLf & sStyles
    
    'If miscTags provided, add them to the inner html
    If miscTags <> "" Then sInnerHTML = sInnerHTML & vbCrLf & miscTags
    
    'return inner html wrapped in header tags
    sHTMLHeader = wrapInHTMLTag(sInnerHTML, "header")
End Function

Function sHtmlStyles(col() As CSS_Style) As String
    'Concatenate all CSS into one long string
    Dim ret As String, css As CSS_Style
    'For Each css In col
    '    ret = ret & cssToString(css) & vbCrLf
    'Next css
    
    For i = LBound(col) To UBound(col)
        css = col(i)
        ret = ret & cssToString(css) & vbCrLf
    Next
    
    'remove trailing vbCrLf
    ret = Left(ret, Len(ret) - Len(vbCrLf))
    
    sHtmlStyles = wrapInHTMLTag(ret, "style")
End Function

Sub testCSSStyles()
    Dim sErrors As CSS_Style
    Dim sWarnings As CSS_Style
    Dim sBasic As CSS_Style
    
    sErrors = newCSSDefault("Errors")
    addCSS sErrors, "background-color", "#ff0000"
    addCSS sErrors, "color", "#ffff00"
    
    sWarnings = newCSSDefault("Warnings")
    addCSS sWarnings, "background-color", "#ffff00"
    addCSS sWarnings, "color", "#000000"
    
    sBasic = newCSSDefault("Basic")
    addCSS sBasic, "color", "#bbbbbb"
    
    Dim CSSStyles(1 To 3) As CSS_Style
    CSSStyles(1) = sErrors
    CSSStyles(2) = sWarnings
    CSSStyles(3) = sBasic
    
    Dim sInnerBody As String
    
    'Demonstrating errors
    On Error GoTo ErrorOccurred
    Dim j, k As Integer: j = 0: k = 10 / j
    GoTo FinishError
ErrorOccurred:
    On Error GoTo 0
    appendInnerHTML sInnerBody, "Error: 10/" & j & " cannot be evaluated", sErrors
FinishError:

    'Demonstrate warnings
    If Cells(1, 1) = "" Then
        appendInnerHTML sInnerBody, "Warning: Blank cell detected in A1 - This could cause issues", sWarnings
    End If
    
    'Demonstrate basic
    appendInnerHTML sInnerBody, "Doing stuff", sBasic
    appendInnerHTML sInnerBody, "Doing more stuff", sBasic
    appendInnerHTML sInnerBody, "Doing other stuff", sBasic
    appendInnerHTML sInnerBody, "Finishing stuff", sBasic
    
    
    Dim sInnerHTML As String
    
    sInnerHTML = sHTMLHeader("Error Log", sHtmlStyles(CSSStyles)) & _
            vbCrLf & wrapInHTMLTag(sInnerBody, "body")
    
    Debug.Print wrapInHTMLTag(sInnerHTML, "html")
End Sub

Sub appendInnerHTML(sInnerBody As String, ByVal sText As String, css As CSS_Style)
    Dim sep As String: If sInnerBody <> "" Then sep = vbCrLf
    sInnerBody = sInnerBody & sep & wrapInHTMLTag(sText, "div", "class=" & css.sName)
End Sub
