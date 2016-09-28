Public Type CSS_Style
    sName As String
    sStyle As String
End Type

Function cssToString(css As CSS_Style) As String
    cssToString = "." & css.sName & " {" & css.sStyle & "}"
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

Function wrapInHTMLTag(ByVal toWrap As String, tag As String)
    wrapInHTMLTag = "<" & tag & ">" & vbCrLf & sTabs(toWrap, 1) & vbCrLf & "</" & tag & ">"
End Function

Function sTabs(ByVal toTab As String, ByVal numTabs As Integer) As String
    Dim tabStr As String: tabStr = String(numTabs, vbTab)
    sTabs = tabStr & Replace(toTab, vbCrLf, vbCrLf & tabStr)
End Function

Function sHtmlStyles(col As Collection) As String
    'Concatenate all CSS into one long string
    Dim ret As String, css As CSS_Style
    For Each css In col
        ret = ret & cssToString(css) & vbCrLf
    Next css
    
    'remove trailing vbCrLf
    ret = Left(ret, Len(ret) - Len(vbCrLf))
    
    sHtmlStyles = wrapInHTMLTag(wrapInHTMLTag(ret, "style"), "header")
End Function

Sub createCSSStyles()
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
    addCSS sBasic, "color", "#eeeeee"
    
    Dim CSSStyles As New Collection
    CSSStyles.Add sErrors, "Errors"
    CSSStyles.Add sWarnings, "Warnings"
    CSSStyles.Add sBasic, "Basic"
    
    Debug.Print sHtmlStyles(CSSStyles)
    
End Sub
