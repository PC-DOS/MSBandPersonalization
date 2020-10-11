Imports Microsoft.Band
Imports Microsoft.Band.Tiles
Imports Microsoft.Band.Admin
Imports Microsoft.Band.Personalization
Module BandTileManager
    Public BandDefaultTiles As New List(Of AdminBandTile)
    Public BandDefaultTilesNameList As New List(Of String)

    Public BandAvailableTiles As New List(Of AdminBandTile)
    Public BandAvailableTilesNameList As New List(Of String)

    Public UserEditedTiles As New List(Of AdminBandTile)
    Public UserEditedTilesNameList As New List(Of String)

    Public BandCurrentStartStrip As StartStrip
    Public Function IsTilePinned(TileToTest As AdminBandTile) As Boolean
        For Each TileItem In UserEditedTiles
            If TileToTest.Id.ToString = TileItem.Id.ToString Then
                Return True
            End If
        Next
        Return False
    End Function
    Public Sub GetBandTileInfo()
        BandDefaultTiles.Clear()
        BandDefaultTilesNameList.Clear()
        BandAvailableTiles.Clear()
        BandAvailableTilesNameList.Clear()
        UserEditedTiles.Clear()
        UserEditedTilesNameList.Clear()

        BandDefaultTiles = BandClient.GetDefaultTilesNoImages()
        For Each BandTileItem In BandDefaultTiles
            BandDefaultTilesNameList.Add(BandTileItem.Name & " [GUID=" & BandTileItem.Id.ToString & "]")
        Next

        BandCurrentStartStrip = BandClient.GetStartStripNoImages()
        For Each BandTileItem In BandCurrentStartStrip
            UserEditedTiles.Add(BandTileItem)
            UserEditedTilesNameList.Add(BandTileItem.Name & " [GUID=" & BandTileItem.Id.ToString & "]")
        Next

        GenerateAvailableTiles()

    End Sub
    Public Sub GenerateAvailableTiles()
        BandAvailableTiles.Clear()
        BandAvailableTilesNameList.Clear()

        For Each BandTileItem In BandDefaultTiles
            If Not IsTilePinned(BandTileItem) Then
                BandAvailableTiles.Add(BandTileItem)
                BandAvailableTilesNameList.Add(BandTileItem.Name & " [GUID=" & BandTileItem.Id.ToString & "]")
            End If
        Next
    End Sub
    Public Sub SetBandTileInfo()
        BandCurrentStartStrip = New StartStrip(UserEditedTiles)
        BandClient.SetStartStrip(BandCurrentStartStrip)
    End Sub
End Module
