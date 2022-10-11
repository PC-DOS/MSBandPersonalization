Imports System.IO
Imports System.Xml
Imports System.Xml.Serialization
Public Module BandThemeManager
    <XmlRoot(ElementName:="BandColor")> Public Class BandThemeColorData
        <XmlElement(ElementName:="Red")> Public Red As Byte
        <XmlElement(ElementName:="Green")> Public Green As Byte
        <XmlElement(ElementName:="Blue")> Public Blue As Byte
        Public Sub New()
            Red = 0
            Green = 0
            Blue = 0
        End Sub
        Public Sub New(R As Byte, G As Byte, B As Byte)
            Red = R
            Green = G
            Blue = B
        End Sub
    End Class
    <XmlRoot(ElementName:="BandTheme")> Public Class BandThemeData
        <XmlElement(ElementName:="BaseColor")> Public BaseColor As New BandThemeColorData
        <XmlElement(ElementName:="HighlightColor")> Public HighlightColor As New BandThemeColorData
        <XmlElement(ElementName:="LowlightColor")> Public LowlightColor As New BandThemeColorData
        <XmlElement(ElementName:="SecondaryColor")> Public SecondaryColor As New BandThemeColorData
        <XmlElement(ElementName:="HighContrastColor")> Public HighContrastColor As New BandThemeColorData
        <XmlElement(ElementName:="MutedColor")> Public MutedColor As New BandThemeColorData
        Public Sub New()
            With BaseColor
                .Red = 0
                .Green = 175
                .Blue = 245
            End With
            With HighlightColor
                .Red = 0
                .Green = 200
                .Blue = 245
            End With
            With LowlightColor
                .Red = 0
                .Green = 150
                .Blue = 245
            End With
            With SecondaryColor
                .Red = 225
                .Green = 225
                .Blue = 225
            End With
            With HighContrastColor
                .Red = 0
                .Green = 200
                .Blue = 245
            End With
            With MutedColor
                .Red = 0
                .Green = 150
                .Blue = 200
            End With
        End Sub
    End Class
    Public Function LoadBandThemeFromFile(ThemeFilePath As String) As BandThemeData
        Dim BandThemeTemp As New BandThemeData
        Dim Serializer As New XmlSerializer(GetType(BandThemeData))
        Dim XmlFileStream As New FileStream(ThemeFilePath, FileMode.OpenOrCreate)
        BandThemeTemp = Serializer.Deserialize(XmlFileStream)
        Return BandThemeTemp
    End Function
    Public Sub SaveBandThemeToFile(ThemeData As BandThemeData, ThemeFilePath As String)
        Dim Serializer As New XmlSerializer(GetType(BandThemeData))
        Dim XmlFileStream As New FileStream(ThemeFilePath, FileMode.Create)
        Serializer.Serialize(XmlFileStream, ThemeData)
        XmlFileStream.Flush()
        XmlFileStream.Close()
    End Sub
End Module
