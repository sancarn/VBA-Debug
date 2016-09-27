Attribute VB_Name = "CSS"
Type CSS_Style
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
