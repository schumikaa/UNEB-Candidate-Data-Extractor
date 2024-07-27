Module functions
    Public Function GetBrushFromHex(hexColor As String) As SolidColorBrush
        Return New SolidColorBrush(CType(ColorConverter.ConvertFromString(hexColor), Color))
    End Function

    Public Function StatusColorHex(status As String) As SolidColorBrush
        Dim colorCode As String
        If status = "Warning" Then
            colorCode = "#ff9103" ' Orange
        ElseIf status = "Ready" Then
            colorCode = "#08e0ff" ' Light Blue
        ElseIf status = "Success" Then
            colorCode = "#03ffe6" ' Light Green
        ElseIf status = "Danger" Then
            colorCode = "#ff0808" ' Red
        ElseIf status = "Info" Then
            colorCode = "#bd08ff" ' Purple
        ElseIf status = "Inactive" Then
            colorCode = "#969696" ' Gray
        ElseIf status = "Dark" Then
            colorCode = "#000000" ' Black
        ElseIf status = "Default" Then
            colorCode = "#3C68E1" ' Navy Blue
        ElseIf status = "Yellow" Then
            colorCode = "#F6F775" ' Light Yellow
        Else
            colorCode = "#263b47" ' Pale Navy Blue
        End If

        Dim color As Color = DirectCast(ColorConverter.ConvertFromString(colorCode), Color)
        Dim statusColor As New SolidColorBrush(color)
        Return statusColor
    End Function
    Public Function TitleCase(ByRef input As String) As String
        Dim cultureInfo As System.Globalization.CultureInfo = System.Globalization.CultureInfo.CurrentCulture
        Dim textInfo As System.Globalization.TextInfo = cultureInfo.TextInfo
        Return textInfo.ToTitleCase(input.ToLower())
    End Function
End Module
