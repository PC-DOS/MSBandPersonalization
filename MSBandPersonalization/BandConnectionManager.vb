Imports Microsoft.Band
Imports Microsoft.Band.Tiles
Imports Microsoft.Band.Admin
Imports Microsoft.Band.Personalization
Module BandConnectionManager
    Public BandClient As ICargoClient
    Public BandClientLimited As IBandClient
    Public CurrentBandClass As BandClass = BandClass.Unknown
End Module
