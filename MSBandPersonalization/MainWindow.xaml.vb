Imports Microsoft.Band
Imports Microsoft.Band.Admin
Imports Microsoft.Band.Personalization
Imports MahApps.Metro.Controls
Imports MahApps.Metro.Controls.Dialogs
Imports System.IO
Imports System.Windows.Media
Imports System.Windows.Media.Imaging
Imports Microsoft.WindowsAPICodePack.Dialogs
Class MainWindow
    Dim BandList As IBandInfo()
    Dim BandClient As ICargoClient
    Dim EmptyList As New List(Of String)
    Dim IsDeviceConnected As Boolean = False
    Dim BandTileList As New List(Of AdminBandTile)
    Dim BandTileNameList As New List(Of String)
    Dim CurrentBandMeTileImage As WriteableBitmap
    Dim CurrentBandTheme As BandTheme
    Private Sub ClearInfo(Optional LockOperationWindows As Boolean = True)
        txtApplicationVersion.Text = ""
        txtBootloaderVersion.Text = ""
        txtDeviceUniqueID.Text = ""
        txtHardwareVersion.Text = ""
        txtPCBID.Text = ""
        txtSerialNumber.Text = ""
        txtUpdaterVersion.Text = ""
        txtOOBEStage.Text = ""
        txtIsOOBECompleted.Text = ""
        txtDeviceName.Text = ""
        imgMeTileImage.Source = Nothing
        btnColorBase.Background = New SolidColorBrush(Color.FromRgb(0, 175, 245))
        btnColorHighContrast.Background = New SolidColorBrush(Color.FromRgb(0, 200, 245))
        btnColorHighlight.Background = New SolidColorBrush(Color.FromRgb(0, 175, 245))
        btnColorLowlight.Background = New SolidColorBrush(Color.FromRgb(0, 150, 245))
        btnColorMuted.Background = New SolidColorBrush(Color.FromRgb(0, 150, 200))
        btnColorSecondary.Background = New SolidColorBrush(Color.FromRgb(225, 225, 225))
        btnColorBase.Tag = Color.FromRgb(0, 175, 245)
        btnColorHighContrast.Tag = Color.FromRgb(0, 200, 245)
        btnColorHighlight.Tag = Color.FromRgb(0, 175, 245)
        btnColorLowlight.Tag = Color.FromRgb(0, 150, 245)
        btnColorMuted.Tag = Color.FromRgb(0, 150, 200)
        btnColorSecondary.Tag = Color.FromRgb(225, 225, 225)
        If LockOperationWindows Then
            LockOperationWindow()
        End If
    End Sub
    Private Sub LockOperationWindow()
        btnFinalizeOOBE.IsEnabled = False
        btnFactoryReset.IsEnabled = False
        btnEnableRetailDemoMode.IsEnabled = False
        btnDisableRetailDemoMode.IsEnabled = False
        btnBrowseMeTileImage.IsEnabled = False
        btnSetMeTileImage.IsEnabled = False
        txtDeviceName.IsEnabled = False
        btnColorBase.IsEnabled = False
        btnColorHighlight.IsEnabled = False
        btnColorLowlight.IsEnabled = False
        btnColorSecondary.IsEnabled = False
        btnColorHighContrast.IsEnabled = False
        btnColorMuted.IsEnabled = False
        btnSetBandTheme.IsEnabled = False
    End Sub
    Private Sub UnockOperationWindow()
        btnFinalizeOOBE.IsEnabled = True
        btnFactoryReset.IsEnabled = True
        btnEnableRetailDemoMode.IsEnabled = True
        btnDisableRetailDemoMode.IsEnabled = True
        btnBrowseMeTileImage.IsEnabled = True
        btnSetMeTileImage.IsEnabled = True
        txtDeviceName.IsEnabled = True
        btnColorBase.IsEnabled = True
        btnColorHighlight.IsEnabled = True
        btnColorLowlight.IsEnabled = True
        btnColorSecondary.IsEnabled = True
        btnColorHighContrast.IsEnabled = True
        btnColorMuted.IsEnabled = True
        btnSetBandTheme.IsEnabled = True
    End Sub
    Private Async Sub OnBandDisconnected(sender As Object, e As EventArgs)
        Await ShowMessageAsync("已斷開連線", "已斷開與當前裝置的連線，請嘗試重新連線。")
        Try
            BandClient.CloseSession()
        Catch ex As Exception

        End Try
        Try
            BandClient.Dispose()
        Catch ex As Exception

        End Try
        IsDeviceConnected = False
        LockOperationWindow()
    End Sub

    Private Async Sub MainWindow_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        ClearInfo()
        Dim IsDeviceFindError As Boolean = False
        Dim ErrorMessage As String = "無法正確描述的例外情況。"
        Try
            BandList = Await BandAdminClientManager.Instance.GetBandsAsync()
            Dim i As Integer
            Dim DeviceNameList As New List(Of String)
            lstDevices.ItemsSource = EmptyList
            If BandList.Length >= 1 Then
                For i = 0 To BandList.Length - 1
                    DeviceNameList.Add(BandList(i).Name & " [通過" & IIf(BandList(i).ConnectionType = BandConnectionType.Bluetooth, "藍牙", " USB ") & "連線]")
                Next
                lstDevices.ItemsSource = DeviceNameList
                lstDevices.SelectedIndex = 0
                btnConnect.IsEnabled = True
            Else
                btnConnect.IsEnabled = False
            End If
        Catch ex As Exception
            IsDeviceFindError = True
            ErrorMessage = ex.Message
            lstDevices.ItemsSource = EmptyList
            btnConnect.IsEnabled = False
        End Try
        If IsDeviceFindError Then
            Await ShowMessageAsync("無法獲取已連線的裝置清單", "試圖獲取已連線的裝置清單時發生例外情況:" & vbCrLf & ErrorMessage)
        End If
    End Sub

    Private Async Sub btnRefreshDevice_Click(sender As Object, e As RoutedEventArgs) Handles btnRefreshDevice.Click
        Dim IsDeviceFindError As Boolean = False
        Dim ErrorMessage As String = "無法正確描述的例外情況。"
        Try
            BandList = Await BandAdminClientManager.Instance.GetBandsAsync()
            Dim i As Integer
            Dim DeviceNameList As New List(Of String)
            lstDevices.ItemsSource = EmptyList
            If BandList.Length >= 1 Then
                For i = 0 To BandList.Length - 1
                    DeviceNameList.Add(BandList(i).Name & " [透過" & IIf(BandList(i).ConnectionType = BandConnectionType.Bluetooth, "藍牙", " USB ") & "連線]")
                Next
                lstDevices.ItemsSource = DeviceNameList
                lstDevices.SelectedIndex = 0
                btnConnect.IsEnabled = True
            Else
                btnConnect.IsEnabled = False
            End If
        Catch ex As Exception
            IsDeviceFindError = True
            ErrorMessage = ex.Message
            lstDevices.ItemsSource = EmptyList
            btnConnect.IsEnabled = False
        End Try
        If IsDeviceFindError Then
            Await ShowMessageAsync("無法獲取已連線的裝置清單", "試圖獲取已連線的裝置清單時發生例外情況:" & vbCrLf & ErrorMessage)
        End If
    End Sub

    Private Async Sub btnConnect_Click(sender As Object, e As RoutedEventArgs) Handles btnConnect.Click
        ClearInfo()
        Dim IsDeviceConnectionError As Boolean = False
        Dim ErrorMessage As String = "無法正確描述的例外情況。"
        Try
            BandClient.CloseSession()
        Catch ex As Exception

        End Try
        IsDeviceConnected = False
        If lstDevices.SelectedIndex < 0 Then
            Await ShowMessageAsync("無法連線到裝置", "無法連線到裝置，因為沒有任何可供連線的裝置。")
        End If
        Try
            BandClient = BandAdminClientManager.Instance.Connect(BandList(lstDevices.SelectedIndex))
        Catch ex As Exception
            IsDeviceConnectionError = True
            ErrorMessage = ex.Message
        End Try
        If IsDeviceConnectionError Then
            IsDeviceConnected = False
            Await ShowMessageAsync("無法連線到裝置", "無法連線到裝置，因為發生例外情況:" & vbCrLf & ErrorMessage)
        Else
            IsDeviceConnected = True
        End If
        If IsDeviceConnected Then
            Try
                txtSerialNumber.Text = BandClient.SerialNumber
            Catch ex As Exception
                txtSerialNumber.Text = "試圖獲取資料時發生例外情況: " & ex.Message
            End Try
            Try
                txtDeviceUniqueID.Text = BandClient.DeviceUniqueId.ToString()
            Catch ex As Exception
                txtDeviceUniqueID.Text = "試圖獲取資料時發生例外情況: " & ex.Message
            End Try
            'Try
            '    txtUserAgent.Text = BandClient.UserAgent
            'Catch ex As Exception
            '    txtUserAgent.Text = "試圖獲取資料時發生例外情況: " & ex.Message
            'End Try
            Try
                txtHardwareVersion.Text = Await BandClient.GetHardwareVersionAsync()
            Catch ex As Exception
                txtHardwareVersion.Text = "試圖獲取資料時發生例外情況: " & ex.Message
            End Try
            Try
                txtApplicationVersion.Text = BandClient.FirmwareVersions.ApplicationVersion.ToString()
            Catch ex As Exception
                txtApplicationVersion.Text = "試圖獲取資料時發生例外情況: " & ex.Message
            End Try
            Try
                txtBootloaderVersion.Text = BandClient.FirmwareVersions.BootloaderVersion.ToString()
            Catch ex As Exception
                txtBootloaderVersion.Text = "試圖獲取資料時發生例外情況: " & ex.Message
            End Try
            Try
                txtUpdaterVersion.Text = BandClient.FirmwareVersions.UpdaterVersion.ToString()
            Catch ex As Exception
                txtUpdaterVersion.Text = "試圖獲取資料時發生例外情況: " & ex.Message
            End Try
            Try
                txtPCBID.Text = BandClient.FirmwareVersions.PcbId.ToString()
            Catch ex As Exception
                txtPCBID.Text = "試圖獲取資料時發生例外情況: " & ex.Message
            End Try
            Try
                txtOOBEStage.Text = BandClient.GetOobeStage().ToString
            Catch ex As Exception
                txtOOBEStage.Text = "試圖獲取資料時發生例外情況: " & ex.Message
            End Try
            Try
                txtIsOOBECompleted.Text = IIf(BandClient.GetDeviceOobeCompleted(), "是", "否")
            Catch ex As Exception
                txtIsOOBECompleted.Text = "試圖獲取資料時發生例外情況: " & ex.Message
            End Try
            Try
                txtDeviceName.Text = BandClient.GetUserProfileFromDevice.DeviceSettings.DeviceName
            Catch ex As Exception
                txtIsOOBECompleted.Text = "試圖獲取資料時發生例外情況: " & ex.Message
            End Try
            'BandTileList.Clear()
            'BandTileNameList.Clear()
            'Try
            '    BandTileList = Await BandClient.GetDefaultTilesAsync()
            '    For Each SingleTile As AdminBandTile In BandTileList
            '        BandTileNameList.Add(SingleTile.Name & " [GUID=" & SingleTile.Id.ToString() & "]")
            '    Next
            'Catch ex As Exception
            '    MessageBox.Show("試圖獲取裝置動態磚資料時發生例外情況: " & ex.Message, "錯誤", MessageBoxButton.OK, MessageBoxImage.Error)
            '    BandTileList.Clear()
            '    BandTileNameList.Clear()
            'End Try
            'lstTiles.ItemsSource = BandTileNameList
            Try
                Dim BandImage As BandImage
                BandImage = Await BandClient.GetMeTileImageAsync()
                Dim BandImageBitmap As WriteableBitmap
                BandImageBitmap = BandImage.ToWriteableBitmap()
                CurrentBandMeTileImage = BandImageBitmap
                imgMeTileImage.Source = BandImageBitmap
                imgMeTileImage.Tag = "CurrentBandMeTileImage"
            Catch ex As Exception
                MessageBox.Show("試圖獲取主動態磚背景圖片時發生例外情況: " & ex.Message, "錯誤", MessageBoxButton.OK, MessageBoxImage.Error)
                imgMeTileImage.Source = Nothing
                CurrentBandMeTileImage = Nothing
                imgMeTileImage.Tag = ""
            End Try
            Try
                CurrentBandTheme = Await BandClient.PersonalizationManager.GetThemeAsync
                btnColorBase.Background = New SolidColorBrush(Color.FromRgb(CurrentBandTheme.Base.R, CurrentBandTheme.Base.G, CurrentBandTheme.Base.B))
                btnColorHighContrast.Background = New SolidColorBrush(Color.FromRgb(CurrentBandTheme.HighContrast.R, CurrentBandTheme.HighContrast.G, CurrentBandTheme.HighContrast.B))
                btnColorHighlight.Background = New SolidColorBrush(Color.FromRgb(CurrentBandTheme.Highlight.R, CurrentBandTheme.Highlight.G, CurrentBandTheme.Highlight.B))
                btnColorLowlight.Background = New SolidColorBrush(Color.FromRgb(CurrentBandTheme.Lowlight.R, CurrentBandTheme.Lowlight.G, CurrentBandTheme.Lowlight.B))
                btnColorMuted.Background = New SolidColorBrush(Color.FromRgb(CurrentBandTheme.Muted.R, CurrentBandTheme.Muted.G, CurrentBandTheme.Muted.B))
                btnColorSecondary.Background = New SolidColorBrush(Color.FromRgb(CurrentBandTheme.SecondaryText.R, CurrentBandTheme.SecondaryText.G, CurrentBandTheme.SecondaryText.B))
                btnColorBase.Tag = Color.FromRgb(CurrentBandTheme.Base.R, CurrentBandTheme.Base.G, CurrentBandTheme.Base.B)
                btnColorHighContrast.Tag = Color.FromRgb(CurrentBandTheme.HighContrast.R, CurrentBandTheme.HighContrast.G, CurrentBandTheme.HighContrast.B)
                btnColorHighlight.Tag = Color.FromRgb(CurrentBandTheme.Highlight.R, CurrentBandTheme.Highlight.G, CurrentBandTheme.Highlight.B)
                btnColorLowlight.Tag = Color.FromRgb(CurrentBandTheme.Lowlight.R, CurrentBandTheme.Lowlight.G, CurrentBandTheme.Lowlight.B)
                btnColorMuted.Tag = Color.FromRgb(CurrentBandTheme.Muted.R, CurrentBandTheme.Muted.G, CurrentBandTheme.Muted.B)
                btnColorSecondary.Tag = Color.FromRgb(CurrentBandTheme.SecondaryText.R, CurrentBandTheme.SecondaryText.G, CurrentBandTheme.SecondaryText.B)
            Catch ex As Exception
                MessageBox.Show("試圖獲取裝置佈景主題時發生例外情況: " & ex.Message, "錯誤", MessageBoxButton.OK, MessageBoxImage.Error)
                btnColorBase.Background = New SolidColorBrush(Color.FromRgb(0, 175, 245))
                btnColorHighContrast.Background = New SolidColorBrush(Color.FromRgb(0, 200, 245))
                btnColorHighlight.Background = New SolidColorBrush(Color.FromRgb(0, 175, 245))
                btnColorLowlight.Background = New SolidColorBrush(Color.FromRgb(0, 150, 245))
                btnColorMuted.Background = New SolidColorBrush(Color.FromRgb(0, 150, 200))
                btnColorSecondary.Background = New SolidColorBrush(Color.FromRgb(225, 225, 225))
                btnColorBase.Tag = Color.FromRgb(0, 175, 245)
                btnColorHighContrast.Tag = Color.FromRgb(0, 200, 245)
                btnColorHighlight.Tag = Color.FromRgb(0, 175, 245)
                btnColorLowlight.Tag = Color.FromRgb(0, 150, 245)
                btnColorMuted.Tag = Color.FromRgb(0, 150, 200)
                btnColorSecondary.Tag = Color.FromRgb(225, 225, 225)
            End Try
            UnockOperationWindow()
        Else
            LockOperationWindow()
        End If
    End Sub

    Private Async Sub btnFinalizeOOBE_Click(sender As Object, e As RoutedEventArgs) Handles btnFinalizeOOBE.Click
        Try
            Await BandClient.SetOobeStageAsync(OobeStage.PressActionButton)
            'Await BandClient.FinalizeOobeAsync()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "錯誤", MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Sub

    Private Async Sub btnFactoryReset_Click(sender As Object, e As RoutedEventArgs) Handles btnFactoryReset.Click
        If MessageBox.Show("您確定要將裝置恢復到原廠設定嗎?", "恢復原廠設定", MessageBoxButton.YesNo, MessageBoxImage.Exclamation) = MessageBoxResult.Yes Then
            Try
                Await BandClient.CargoSystemSettingsFactoryResetAsync()
            Catch ex As Exception
                MessageBox.Show(ex.Message, "錯誤", MessageBoxButton.OK, MessageBoxImage.Error)
            End Try
        End If
    End Sub

    Private Async Sub btnEnableRetailDemoMode_Click(sender As Object, e As RoutedEventArgs) Handles btnEnableRetailDemoMode.Click
        If MessageBox.Show("您確定要啟用示範模式嗎?", "啟用示範模式", MessageBoxButton.YesNo, MessageBoxImage.Question) = MessageBoxResult.Yes Then
            Try
                Await BandClient.EnableRetailDemoModeAsync()
            Catch ex As Exception
                MessageBox.Show(ex.Message, "錯誤", MessageBoxButton.OK, MessageBoxImage.Error)
            End Try
        End If
    End Sub

    Private Async Sub btnDisableRetailDemoMode_Click(sender As Object, e As RoutedEventArgs) Handles btnDisableRetailDemoMode.Click
        If MessageBox.Show("您確定要停用示範模式嗎?", "停用示範模式", MessageBoxButton.YesNo, MessageBoxImage.Question) = MessageBoxResult.Yes Then
            Try
                Await BandClient.DisableRetailDemoModeAsync()
            Catch ex As Exception
                MessageBox.Show(ex.Message, "錯誤", MessageBoxButton.OK, MessageBoxImage.Error)
            End Try
        End If
    End Sub

    Private Sub txtDeviceName_MouseUp(sender As Object, e As MouseButtonEventArgs) Handles txtDeviceName.MouseUp
        Dim NewBandName As String
        NewBandName = InputBox("請輸入新的裝置名稱。", "重新命名裝置", txtDeviceName.Text)
        If NewBandName.Trim() <> "" Then
            Try
                Dim NewBandUserProfile As IUserProfile
                NewBandUserProfile = BandClient.GetUserProfileFromDevice
                NewBandUserProfile.DeviceSettings.DeviceName = NewBandName
                BandClient.SaveUserProfileToBandOnly(NewBandUserProfile)
            Catch ex As Exception
                MessageBox.Show("試圖重新命名裝置時發生例外情況: " & ex.Message)
            End Try
            Try
                txtDeviceName.Text = BandClient.GetUserProfileFromDevice.DeviceSettings.DeviceName
            Catch ex As Exception
                txtIsOOBECompleted.Text = "試圖獲取資料時發生例外情況: " & ex.Message
            End Try
        End If
    End Sub

    Private Sub btnBrowseMeTileImage_Click(sender As Object, e As RoutedEventArgs) Handles btnBrowseMeTileImage.Click
        Dim ImageBrowseDialog As New CommonOpenFileDialog
        With ImageBrowseDialog
            .Filters.Add(New CommonFileDialogFilter("圖片", ".jpg;.jpeg;.bmp;.bip;.png;.gif"))
            .Filters.Add(New CommonFileDialogFilter("所有檔案", ".*"))
        End With
        If ImageBrowseDialog.ShowDialog = CommonFileDialogResult.Ok Then
            Try
                imgMeTileImage.Source = New BitmapImage(New Uri(ImageBrowseDialog.FileName))
                imgMeTileImage.Tag = ImageBrowseDialog.FileName
            Catch ex As Exception
                MessageBox.Show("打開圖片""" & ImageBrowseDialog.FileName & """時發生例外情況: " & ex.Message, "錯誤", MessageBoxButton.OK, MessageBoxImage.Error)
                imgMeTileImage.Source = CurrentBandMeTileImage
            End Try
        End If
    End Sub

    Private Async Sub btnSetMeTileImage_Click(sender As Object, e As RoutedEventArgs) Handles btnSetMeTileImage.Click
        'Code adapted from unBand
        Try
            Dim BitmapSource As New BitmapImage()
            'BitmapSource.BeginInit()
            BitmapSource = imgMeTileImage.Source
            BitmapSource.DecodePixelHeight = 102
            BitmapSource.DecodePixelWidth = 310
            'BitmapSource.EndInit()
            Dim Pbgra32Image = New FormatConvertedBitmap(BitmapSource, PixelFormats.Pbgra32, Nothing, 0)
            Dim NewMeTileImage As New WriteableBitmap(Pbgra32Image)
            Dim NewBandMeTileImage As BandImage = NewMeTileImage.ToBandImage()
            Await BandClient.PersonalizationManager.SetMeTileImageAsync(NewBandMeTileImage)
        Catch ex As Exception
            MessageBox.Show("試圖設置主動態磚背景圖片時發生例外情況: " & ex.Message, "錯誤", MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
        Try
            Dim BandImage As BandImage
            BandImage = Await BandClient.GetMeTileImageAsync()
            Dim BandImageBitmap As WriteableBitmap
            BandImageBitmap = BandImage.ToWriteableBitmap()
            CurrentBandMeTileImage = BandImageBitmap
            imgMeTileImage.Source = BandImageBitmap
            imgMeTileImage.Tag = "CurrentBandMeTileImage"
        Catch ex As Exception
            MessageBox.Show("試圖獲取主動態磚背景圖片時發生例外情況: " & ex.Message, "錯誤", MessageBoxButton.OK, MessageBoxImage.Error)
            imgMeTileImage.Tag = ""
        End Try
    End Sub

    Private Sub btnColorBase_MouseUp(sender As Object, e As MouseButtonEventArgs) Handles btnColorBase.MouseUp
        Dim ColorPicker As New Forms.ColorDialog
        Dim CurrentColor As Color = btnColorBase.Tag
        Dim CurrentSelectedColor As System.Drawing.Color = System.Drawing.Color.FromArgb(CurrentColor.A, CurrentColor.R, CurrentColor.G, CurrentColor.B)
        With ColorPicker
            .AllowFullOpen = True
            .SolidColorOnly = True
            .Color = CurrentSelectedColor
        End With
        If ColorPicker.ShowDialog() = Forms.DialogResult.OK Then
            Dim SelectedColor As Color = Color.FromRgb(ColorPicker.Color.R, ColorPicker.Color.G, ColorPicker.Color.B)
            btnColorBase.Background = New SolidColorBrush(SelectedColor)
            btnColorBase.Tag = SelectedColor
        End If
    End Sub

    Private Sub btnColorHighContrast_MouseUp(sender As Object, e As MouseButtonEventArgs) Handles btnColorHighContrast.MouseUp
        Dim ColorPicker As New Forms.ColorDialog
        Dim CurrentColor As Color = btnColorHighContrast.Tag
        Dim CurrentSelectedColor As System.Drawing.Color = System.Drawing.Color.FromArgb(CurrentColor.A, CurrentColor.R, CurrentColor.G, CurrentColor.B)
        With ColorPicker
            .AllowFullOpen = True
            .SolidColorOnly = True
            .Color = CurrentSelectedColor
        End With
        If ColorPicker.ShowDialog() = Forms.DialogResult.OK Then
            Dim SelectedColor As Color = Color.FromRgb(ColorPicker.Color.R, ColorPicker.Color.G, ColorPicker.Color.B)
            btnColorHighContrast.Background = New SolidColorBrush(SelectedColor)
            btnColorHighContrast.Tag = SelectedColor
        End If
    End Sub

    Private Sub btnColorHighlight_MouseUp(sender As Object, e As MouseButtonEventArgs) Handles btnColorHighlight.MouseUp
        Dim ColorPicker As New Forms.ColorDialog
        Dim CurrentColor As Color = btnColorHighlight.Tag
        Dim CurrentSelectedColor As System.Drawing.Color = System.Drawing.Color.FromArgb(CurrentColor.A, CurrentColor.R, CurrentColor.G, CurrentColor.B)
        With ColorPicker
            .AllowFullOpen = True
            .SolidColorOnly = True
            .Color = CurrentSelectedColor
        End With
        If ColorPicker.ShowDialog() = Forms.DialogResult.OK Then
            Dim SelectedColor As Color = Color.FromRgb(ColorPicker.Color.R, ColorPicker.Color.G, ColorPicker.Color.B)
            btnColorHighlight.Background = New SolidColorBrush(SelectedColor)
            btnColorHighlight.Tag = SelectedColor
        End If
    End Sub

    Private Sub btnColorLowlight_MouseUp(sender As Object, e As MouseButtonEventArgs) Handles btnColorLowlight.MouseUp
        Dim ColorPicker As New Forms.ColorDialog
        Dim CurrentColor As Color = btnColorLowlight.Tag
        Dim CurrentSelectedColor As System.Drawing.Color = System.Drawing.Color.FromArgb(CurrentColor.A, CurrentColor.R, CurrentColor.G, CurrentColor.B)
        With ColorPicker
            .AllowFullOpen = True
            .SolidColorOnly = True
            .Color = CurrentSelectedColor
        End With
        If ColorPicker.ShowDialog() = Forms.DialogResult.OK Then
            Dim SelectedColor As Color = Color.FromRgb(ColorPicker.Color.R, ColorPicker.Color.G, ColorPicker.Color.B)
            btnColorLowlight.Background = New SolidColorBrush(SelectedColor)
            btnColorLowlight.Tag = SelectedColor
        End If
    End Sub

    Private Sub btnColorMuted_MouseUp(sender As Object, e As MouseButtonEventArgs) Handles btnColorMuted.MouseUp
        Dim ColorPicker As New Forms.ColorDialog
        Dim CurrentColor As Color = btnColorMuted.Tag
        Dim CurrentSelectedColor As System.Drawing.Color = System.Drawing.Color.FromArgb(CurrentColor.A, CurrentColor.R, CurrentColor.G, CurrentColor.B)
        With ColorPicker
            .AllowFullOpen = True
            .SolidColorOnly = True
            .Color = CurrentSelectedColor
        End With
        If ColorPicker.ShowDialog() = Forms.DialogResult.OK Then
            Dim SelectedColor As Color = Color.FromRgb(ColorPicker.Color.R, ColorPicker.Color.G, ColorPicker.Color.B)
            btnColorMuted.Background = New SolidColorBrush(SelectedColor)
            btnColorMuted.Tag = SelectedColor
        End If
    End Sub

    Private Sub btnColorSecondary_MouseUp(sender As Object, e As MouseButtonEventArgs) Handles btnColorSecondary.MouseUp
        Dim ColorPicker As New Forms.ColorDialog
        Dim CurrentColor As Color = btnColorSecondary.Tag
        Dim CurrentSelectedColor As System.Drawing.Color = System.Drawing.Color.FromArgb(CurrentColor.A, CurrentColor.R, CurrentColor.G, CurrentColor.B)
        With ColorPicker
            .AllowFullOpen = True
            .SolidColorOnly = True
            .Color = CurrentSelectedColor
        End With
        If ColorPicker.ShowDialog() = Forms.DialogResult.OK Then
            Dim SelectedColor As Color = Color.FromRgb(ColorPicker.Color.R, ColorPicker.Color.G, ColorPicker.Color.B)
            btnColorSecondary.Background = New SolidColorBrush(SelectedColor)
            btnColorSecondary.Tag = SelectedColor
        End If
    End Sub

    Private Async Sub btnSetBandTheme_Click(sender As Object, e As RoutedEventArgs) Handles btnSetBandTheme.Click
        Try
            With CurrentBandTheme
                Dim CurrentProcessingColor As Color
                CurrentProcessingColor = btnColorBase.Tag
                .Base = CurrentProcessingColor.ToBandColor()
                CurrentProcessingColor = btnColorHighContrast.Tag
                .HighContrast = CurrentProcessingColor.ToBandColor()
                CurrentProcessingColor = btnColorHighlight.Tag
                .Highlight = CurrentProcessingColor.ToBandColor()
                CurrentProcessingColor = btnColorLowlight.Tag
                .Lowlight = CurrentProcessingColor.ToBandColor()
                CurrentProcessingColor = btnColorMuted.Tag
                .Muted = CurrentProcessingColor.ToBandColor()
                CurrentProcessingColor = btnColorSecondary.Tag
                .SecondaryText = CurrentProcessingColor.ToBandColor()
            End With
            Await BandClient.PersonalizationManager.SetThemeAsync(CurrentBandTheme)
        Catch ex As Exception
            MessageBox.Show("試圖設置裝置佈景主題時發生例外情況: " & ex.Message, "錯誤", MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
        Try
            CurrentBandTheme = Await BandClient.PersonalizationManager.GetThemeAsync
            btnColorBase.Background = New SolidColorBrush(Color.FromRgb(CurrentBandTheme.Base.R, CurrentBandTheme.Base.G, CurrentBandTheme.Base.B))
            btnColorHighContrast.Background = New SolidColorBrush(Color.FromRgb(CurrentBandTheme.HighContrast.R, CurrentBandTheme.HighContrast.G, CurrentBandTheme.HighContrast.B))
            btnColorHighlight.Background = New SolidColorBrush(Color.FromRgb(CurrentBandTheme.Highlight.R, CurrentBandTheme.Highlight.G, CurrentBandTheme.Highlight.B))
            btnColorLowlight.Background = New SolidColorBrush(Color.FromRgb(CurrentBandTheme.Lowlight.R, CurrentBandTheme.Lowlight.G, CurrentBandTheme.Lowlight.B))
            btnColorMuted.Background = New SolidColorBrush(Color.FromRgb(CurrentBandTheme.Muted.R, CurrentBandTheme.Muted.G, CurrentBandTheme.Muted.B))
            btnColorSecondary.Background = New SolidColorBrush(Color.FromRgb(CurrentBandTheme.SecondaryText.R, CurrentBandTheme.SecondaryText.G, CurrentBandTheme.SecondaryText.B))
            btnColorBase.Tag = Color.FromRgb(CurrentBandTheme.Base.R, CurrentBandTheme.Base.G, CurrentBandTheme.Base.B)
            btnColorHighContrast.Tag = Color.FromRgb(CurrentBandTheme.HighContrast.R, CurrentBandTheme.HighContrast.G, CurrentBandTheme.HighContrast.B)
            btnColorHighlight.Tag = Color.FromRgb(CurrentBandTheme.Highlight.R, CurrentBandTheme.Highlight.G, CurrentBandTheme.Highlight.B)
            btnColorLowlight.Tag = Color.FromRgb(CurrentBandTheme.Lowlight.R, CurrentBandTheme.Lowlight.G, CurrentBandTheme.Lowlight.B)
            btnColorMuted.Tag = Color.FromRgb(CurrentBandTheme.Muted.R, CurrentBandTheme.Muted.G, CurrentBandTheme.Muted.B)
            btnColorSecondary.Tag = Color.FromRgb(CurrentBandTheme.SecondaryText.R, CurrentBandTheme.SecondaryText.G, CurrentBandTheme.SecondaryText.B)
        Catch ex As Exception
            MessageBox.Show("試圖獲取裝置佈景主題時發生例外情況: " & ex.Message, "錯誤", MessageBoxButton.OK, MessageBoxImage.Error)
            btnColorBase.Background = New SolidColorBrush(Color.FromRgb(0, 175, 245))
            btnColorHighContrast.Background = New SolidColorBrush(Color.FromRgb(0, 200, 245))
            btnColorHighlight.Background = New SolidColorBrush(Color.FromRgb(0, 175, 245))
            btnColorLowlight.Background = New SolidColorBrush(Color.FromRgb(0, 150, 245))
            btnColorMuted.Background = New SolidColorBrush(Color.FromRgb(0, 150, 200))
            btnColorSecondary.Background = New SolidColorBrush(Color.FromRgb(225, 225, 225))
            btnColorBase.Tag = Color.FromRgb(0, 175, 245)
            btnColorHighContrast.Tag = Color.FromRgb(0, 200, 245)
            btnColorHighlight.Tag = Color.FromRgb(0, 175, 245)
            btnColorLowlight.Tag = Color.FromRgb(0, 150, 245)
            btnColorMuted.Tag = Color.FromRgb(0, 150, 200)
            btnColorSecondary.Tag = Color.FromRgb(225, 225, 225)
        End Try
    End Sub
End Class
