Imports Microsoft.Band
Imports Microsoft.Band.Tiles
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
    Dim EmptyList As New List(Of String)
    Dim IsDeviceConnected As Boolean = False
    Dim BandTileList As BandTile()
    Dim BandTileNameList As New List(Of String)
    Dim CurrentBandMeTileImage As WriteableBitmap
    Dim CurrentBandTheme As BandTheme
    Private Sub ClearInfo(Optional LockOperationWindows As Boolean = True)
        txtApplicationVersion.Text = ""
        txtBandClass.Text = ""
        txtBootloaderVersion.Text = ""
        txtDeviceUniqueID.Text = ""
        txtHardwareVersion.Text = ""
        txtPCBID.Text = ""
        txtSerialNumber.Text = ""
        txtUpdaterVersion.Text = ""
        txtOOBEStage.Text = ""
        txtIsOOBECompleted.Text = ""
        txtDeviceName.Text = ""
        txtUserFirstName.Text = ""
        txtUserLastName.Text = ""
        txtUserHeight.Text = ""
        txtUserWeight.Text = ""
        txtEmailAddress.Text = ""
        txtSmsAddress.Text = ""
        txtDateSeparator.Text = ""
        txtDecimalSeparator.Text = ""
        txtNumberSeparator.Text = ""
        cmbDateFormat.SelectedIndex = 0
        cmbTimeFormat.SelectedIndex = 0
        rdbMale.IsChecked = False
        rdbFemale.IsChecked = False
        rdbMetricDistUnit.IsChecked = False
        rdbImperialDistUnit.IsChecked = False
        rdbMetricMassUnit.IsChecked = False
        rdbImperialMassUnit.IsChecked = False
        rdbMetricVolUnit.IsChecked = False
        rdbImperialVolUnit.IsChecked = False
        rdbMetricTempUnit.IsChecked = False
        rdbImperialTempUnit.IsChecked = False
        rdbMetricEnergyUnit.IsChecked = False
        rdbImperialEnergyUnit.IsChecked = False
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
        lstAvailableTiles.ItemsSource = EmptyList
        lstPinnedTiles.ItemsSource = EmptyList
        If LockOperationWindows Then
            LockOperationWindow()
        End If
    End Sub
    Private Sub LockOperationWindow()
        btnFinalizeOOBE.IsEnabled = False
        btnFactoryReset.IsEnabled = False
        btnEnableRetailDemoMode.IsEnabled = False
        btnDisableRetailDemoMode.IsEnabled = False
        btnVibrate.IsEnabled = False
        btnBrowseMeTileImage.IsEnabled = False
        btnSetMeTileImage.IsEnabled = False
        txtDeviceName.IsEnabled = False
        txtUserFirstName.IsEnabled = False
        txtUserLastName.IsEnabled = False
        txtUserHeight.IsEnabled = False
        txtUserWeight.IsEnabled = False
        txtEmailAddress.IsEnabled = False
        txtSmsAddress.IsEnabled = False
        rdbMale.IsEnabled = False
        rdbFemale.IsEnabled = False
        rdbMetricDistUnit.IsEnabled = False
        rdbImperialDistUnit.IsEnabled = False
        rdbMetricMassUnit.IsEnabled = False
        rdbImperialMassUnit.IsEnabled = False
        rdbMetricVolUnit.IsEnabled = False
        rdbImperialVolUnit.IsEnabled = False
        rdbMetricTempUnit.IsEnabled = False
        rdbImperialTempUnit.IsEnabled = False
        rdbMetricEnergyUnit.IsEnabled = False
        rdbImperialEnergyUnit.IsEnabled = False
        btnColorBase.IsEnabled = False
        btnColorHighlight.IsEnabled = False
        btnColorLowlight.IsEnabled = False
        btnColorSecondary.IsEnabled = False
        btnColorHighContrast.IsEnabled = False
        btnColorMuted.IsEnabled = False
        btnSetBandTheme.IsEnabled = False
        chkStepsEnabled.IsEnabled = False
        chkCaloriesEnabled.IsEnabled = False
        chkDistanceEnabled.IsEnabled = False
        txtStepsGoal.IsEnabled = False
        txtCaloriesGoal.IsEnabled = False
        txtDistanceGoal.IsEnabled = False
        btnSetGoal.IsEnabled = False
        lstAvailableTiles.SelectedIndex = -1
        lstAvailableTiles.IsEnabled = False
        lstPinnedTiles.SelectedIndex = -1
        lstPinnedTiles.IsEnabled = False
        btnAddTile.IsEnabled = False
        btnRemoveTile.IsEnabled = False
        btnTileMoveUp.IsEnabled = False
        btnTileMoveDown.IsEnabled = False
        btnSetBandStartStrip.IsEnabled = False
        txtDateSeparator.IsEnabled = False
        txtDecimalSeparator.IsEnabled = False
        txtNumberSeparator.IsEnabled = False
        cmbDateFormat.IsEnabled = False
        cmbTimeFormat.IsEnabled = False
    End Sub
    Private Sub UnockOperationWindow()
        btnFinalizeOOBE.IsEnabled = True
        btnFactoryReset.IsEnabled = True
        btnEnableRetailDemoMode.IsEnabled = True
        btnDisableRetailDemoMode.IsEnabled = True
        btnVibrate.IsEnabled = True
        btnBrowseMeTileImage.IsEnabled = True
        btnSetMeTileImage.IsEnabled = True
        txtDeviceName.IsEnabled = True
        txtUserFirstName.IsEnabled = True
        txtUserLastName.IsEnabled = True
        txtUserHeight.IsEnabled = True
        txtUserWeight.IsEnabled = True
        txtEmailAddress.IsEnabled = True
        txtSmsAddress.IsEnabled = True
        rdbMale.IsEnabled = True
        rdbFemale.IsEnabled = True
        rdbMetricDistUnit.IsEnabled = True
        rdbImperialDistUnit.IsEnabled = True
        rdbMetricMassUnit.IsEnabled = True
        rdbImperialMassUnit.IsEnabled = True
        rdbMetricVolUnit.IsEnabled = True
        rdbImperialVolUnit.IsEnabled = True
        rdbMetricTempUnit.IsEnabled = True
        rdbImperialTempUnit.IsEnabled = True
        rdbMetricEnergyUnit.IsEnabled = True
        rdbImperialEnergyUnit.IsEnabled = True
        btnColorBase.IsEnabled = True
        btnColorHighlight.IsEnabled = True
        btnColorLowlight.IsEnabled = True
        btnColorSecondary.IsEnabled = True
        btnColorHighContrast.IsEnabled = True
        btnColorMuted.IsEnabled = True
        btnSetBandTheme.IsEnabled = True
        chkStepsEnabled.IsEnabled = True
        chkCaloriesEnabled.IsEnabled = True
        chkDistanceEnabled.IsEnabled = True
        txtStepsGoal.IsEnabled = chkStepsEnabled.IsChecked
        txtCaloriesGoal.IsEnabled = chkCaloriesEnabled.IsChecked
        txtDistanceGoal.IsEnabled = chkDistanceEnabled.IsChecked
        btnSetGoal.IsEnabled = True
        lstAvailableTiles.IsEnabled = True
        lstPinnedTiles.IsEnabled = True
        btnSetBandStartStrip.IsEnabled = True
        lstAvailableTiles.SelectedIndex = -1
        lstPinnedTiles.SelectedIndex = -1
        txtDateSeparator.IsEnabled = True
        txtDecimalSeparator.IsEnabled = True
        txtNumberSeparator.IsEnabled = True
        cmbDateFormat.IsEnabled = True
        cmbTimeFormat.IsEnabled = True
    End Sub
    Private Sub RefreshTileList(Optional NextAvailableTilesListOriginalSelectedIndex As Integer = -2, Optional NextPinnedTilesListOriginalSelectedIndex As Integer = -2)
        Dim AvailableTilesListOriginalSelectedIndex As Integer = NextAvailableTilesListOriginalSelectedIndex
        Dim PinnedTilesListOriginalSelectedIndex As Integer = NextPinnedTilesListOriginalSelectedIndex
        If NextAvailableTilesListOriginalSelectedIndex = -2 Then
            AvailableTilesListOriginalSelectedIndex = lstAvailableTiles.SelectedIndex
        End If
        If NextPinnedTilesListOriginalSelectedIndex = -2 Then
            PinnedTilesListOriginalSelectedIndex = lstPinnedTiles.SelectedIndex
        End If

        lstAvailableTiles.ItemsSource = EmptyList
        lstPinnedTiles.ItemsSource = EmptyList

        lstAvailableTiles.ItemsSource = BandAvailableTilesNameList
        lstPinnedTiles.ItemsSource = UserEditedTilesNameList

        Try
            lstAvailableTiles.SelectedIndex = AvailableTilesListOriginalSelectedIndex
        Catch ex As Exception
            lstAvailableTiles.SelectedIndex = -1
        End Try
        Try
            lstPinnedTiles.SelectedIndex = PinnedTilesListOriginalSelectedIndex
        Catch ex As Exception
            lstPinnedTiles.SelectedIndex = -1
        End Try

        If lstPinnedTiles.SelectedIndex > -1 Then
            btnRemoveTile.IsEnabled = True
            btnTileMoveUp.IsEnabled = True
            btnTileMoveDown.IsEnabled = True
            If lstPinnedTiles.SelectedIndex = 0 Then
                btnTileMoveUp.IsEnabled = False
            End If
            If lstPinnedTiles.SelectedIndex = UserEditedTilesNameList.Count - 1 Then
                btnTileMoveDown.IsEnabled = False
            End If
        End If
        If lstAvailableTiles.SelectedIndex > -1 Then
            btnAddTile.IsEnabled = True
        End If

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
        CurrentBandClass = BandClass.Unknown
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
        CurrentBandClass = BandClass.Unknown
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
                CurrentBandClass = BandClient.ConnectedBandConstants.BandClass
                Select Case CurrentBandClass
                    Case BandClass.Cargo
                        txtBandClass.Text = "Microsoft Band (Cargo)"
                    Case BandClass.Envoy
                        txtBandClass.Text = "Microsoft Band 2 (Envoy)"
                    Case Else
                        txtBandClass.Text = "不明"
                End Select
            Catch ex As Exception
                txtBandClass.Text = "試圖獲取資料時發生例外情況: " & ex.Message
                CurrentBandClass = BandClass.Unknown
            End Try
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
            'BEGIN
            'GetUserProfileFromDevice()方法不支援Microsoft Band 2
            If CurrentBandClass = BandClass.Cargo Then
                Try
                    txtDeviceName.Text = BandClient.GetUserProfileFromDevice.DeviceSettings.DeviceName
                Catch ex As Exception
                    txtDeviceName.Text = "試圖獲取資料時發生例外情況: " & ex.Message
                End Try
                Try
                    txtUserFirstName.Text = BandClient.GetUserProfileFromDevice.FirstName
                Catch ex As Exception
                    txtUserFirstName.Text = "試圖獲取資料時發生例外情況: " & ex.Message
                End Try
                Try
                    txtUserLastName.Text = BandClient.GetUserProfileFromDevice.LastName
                Catch ex As Exception
                    txtUserLastName.Text = "試圖獲取資料時發生例外情況: " & ex.Message
                End Try
                Try
                    txtUserHeight.Text = BandClient.GetUserProfileFromDevice.Height
                Catch ex As Exception
                    txtUserHeight.Text = "試圖獲取資料時發生例外情況: " & ex.Message
                End Try
                Try
                    txtUserWeight.Text = BandClient.GetUserProfileFromDevice.Weight
                Catch ex As Exception
                    txtUserWeight.Text = "試圖獲取資料時發生例外情況: " & ex.Message
                End Try
                Try
                    txtEmailAddress.Text = BandClient.GetUserProfileFromDevice.EmailAddress
                Catch ex As Exception
                    txtEmailAddress.Text = "試圖獲取資料時發生例外情況: " & ex.Message
                End Try
                Try
                    txtSmsAddress.Text = BandClient.GetUserProfileFromDevice.SmsAddress
                Catch ex As Exception
                    txtSmsAddress.Text = "試圖獲取資料時發生例外情況: " & ex.Message
                End Try
                Try
                    If BandClient.GetUserProfileFromDevice.Gender = Gender.Male Then
                        rdbMale.IsChecked = True
                        rdbFemale.IsChecked = False
                    Else
                        rdbMale.IsChecked = False
                        rdbFemale.IsChecked = True
                    End If
                Catch ex As Exception
                    rdbMale.IsChecked = False
                    rdbFemale.IsChecked = False
                End Try
                Try
                    If BandClient.GetUserProfileFromDevice.DeviceSettings.LocaleSettings.DistanceLongUnits = DistanceUnitType.Metric Then
                        rdbMetricDistUnit.IsChecked = True
                        rdbImperialDistUnit.IsChecked = False
                    Else
                        rdbMetricDistUnit.IsChecked = False
                        rdbImperialDistUnit.IsChecked = True
                    End If
                Catch ex As Exception
                    rdbMetricDistUnit.IsChecked = False
                    rdbImperialDistUnit.IsChecked = False
                End Try
                Try
                    If BandClient.GetUserProfileFromDevice.DeviceSettings.LocaleSettings.MassUnits = MassUnitType.Metric Then
                        rdbMetricMassUnit.IsChecked = True
                        rdbImperialMassUnit.IsChecked = False
                    Else
                        rdbMetricMassUnit.IsChecked = False
                        rdbImperialMassUnit.IsChecked = True
                    End If
                Catch ex As Exception
                    rdbMetricMassUnit.IsChecked = False
                    rdbImperialMassUnit.IsChecked = False
                End Try
                Try
                    If BandClient.GetUserProfileFromDevice.DeviceSettings.LocaleSettings.VolumeUnits = VolumeUnitType.Metric Then
                        rdbMetricVolUnit.IsChecked = True
                        rdbImperialVolUnit.IsChecked = False
                    Else
                        rdbMetricVolUnit.IsChecked = False
                        rdbImperialVolUnit.IsChecked = True
                    End If
                Catch ex As Exception
                    rdbMetricVolUnit.IsChecked = False
                    rdbImperialVolUnit.IsChecked = False
                End Try
                Try
                    If BandClient.GetUserProfileFromDevice.DeviceSettings.LocaleSettings.TemperatureUnits = TemperatureUnitType.Metric Then
                        rdbMetricTempUnit.IsChecked = True
                        rdbImperialTempUnit.IsChecked = False
                    Else
                        rdbMetricTempUnit.IsChecked = False
                        rdbImperialTempUnit.IsChecked = True
                    End If
                Catch ex As Exception
                    rdbMetricTempUnit.IsChecked = False
                    rdbImperialTempUnit.IsChecked = False
                End Try
                Try
                    If BandClient.GetUserProfileFromDevice.DeviceSettings.LocaleSettings.EnergyUnits = EnergyUnitType.Metric Then
                        rdbMetricEnergyUnit.IsChecked = True
                        rdbImperialEnergyUnit.IsChecked = False
                    Else
                        rdbMetricEnergyUnit.IsChecked = False
                        rdbImperialEnergyUnit.IsChecked = True
                    End If
                Catch ex As Exception
                    rdbMetricEnergyUnit.IsChecked = False
                    rdbImperialEnergyUnit.IsChecked = False
                End Try
                Try
                    cmbDateFormat.SelectedIndex = BandClient.GetUserProfileFromDevice.DeviceSettings.LocaleSettings.DateFormat
                Catch ex As Exception
                    cmbDateFormat.SelectedIndex = 0
                End Try
                Try
                    cmbTimeFormat.SelectedIndex = BandClient.GetUserProfileFromDevice.DeviceSettings.LocaleSettings.TimeFormat
                Catch ex As Exception
                    cmbTimeFormat.SelectedIndex = 0
                End Try
                Try
                    txtNumberSeparator.Text = BandClient.GetUserProfileFromDevice.DeviceSettings.LocaleSettings.NumberSeparator
                Catch ex As Exception
                    txtNumberSeparator.Text = "試圖獲取小數符號時發生例外情況: " & ex.Message
                End Try
                Try
                    txtDecimalSeparator.Text = BandClient.GetUserProfileFromDevice.DeviceSettings.LocaleSettings.DecimalSeparator
                Catch ex As Exception
                    txtDecimalSeparator.Text = "試圖獲取數字分位符號時發生例外情況: " & ex.Message
                End Try
                Try
                    txtDateSeparator.Text = BandClient.GetUserProfileFromDevice.DeviceSettings.LocaleSettings.DateSeparator
                Catch ex As Exception
                    txtDateSeparator.Text = "試圖獲取日期分位符號時發生例外情況: " & ex.Message
                End Try
                stkCargoOnly.Visibility = Windows.Visibility.Visible
            Else
                stkCargoOnly.Visibility = Windows.Visibility.Collapsed
            End If
            '支援Microsoft Band 2後移除此因應措施
            'END
            Try
                GetBandTileInfo()
                lstAvailableTiles.ItemsSource = BandAvailableTilesNameList
                lstPinnedTiles.ItemsSource = UserEditedTilesNameList
            Catch ex As Exception
                MessageBox.Show("試圖獲取裝置動態磚資料時發生例外情況: " & ex.Message, "錯誤", MessageBoxButton.OK, MessageBoxImage.Error)
                lstAvailableTiles.ItemsSource = EmptyList
                lstPinnedTiles.ItemsSource = EmptyList
            End Try
            Try
                Dim BandImage As BandImage
                BandImage = Await BandClient.PersonalizationManager.GetMeTileImageAsync()
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
                CurrentBandTheme = Await BandClient.PersonalizationManager.GetThemeAsync()
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
            If CurrentBandClass <> BandClass.Cargo Then
                Await ShowMessageAsync("部分功能已停用", "部分設定項目目前僅適用於第一代 Microsoft Band，而您似乎連線到了一個更新的 Microsoft Band 裝置，因此這些設定項目已被停用。")
            End If
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
                MessageBox.Show("試圖重新命名裝置時發生例外情況: " & ex.Message, "錯誤", MessageBoxButton.OK, MessageBoxImage.Error)
            End Try
            Try
                txtDeviceName.Text = BandClient.GetUserProfileFromDevice.DeviceSettings.DeviceName
            Catch ex As Exception
                txtDeviceName.Text = "試圖獲取資料時發生例外情況: " & ex.Message
            End Try
        End If
    End Sub

    Private Sub btnBrowseMeTileImage_Click(sender As Object, e As RoutedEventArgs) Handles btnBrowseMeTileImage.Click
        Dim ImageBrowseDialog As New CommonOpenFileDialog
        With ImageBrowseDialog
            .EnsureFileExists = True
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
            'BitmapSource.DecodePixelHeight = 102
            'BitmapSource.DecodePixelWidth = 310
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
            BandImage = Await BandClient.PersonalizationManager.GetMeTileImageAsync()
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
            CurrentBandTheme = Await BandClient.PersonalizationManager.GetThemeAsync()
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

    Private Sub btnSetGoal_Click(sender As Object, e As RoutedEventArgs) Handles btnSetGoal.Click
        Dim StepCountGoal As Integer
        Dim CaloriesGoal As Integer
        Dim DistanceGoal As Integer
        Try
            StepCountGoal = txtStepsGoal.Text
        Catch ex As Exception
            StepCountGoal = 5000
        End Try
        Try
            CaloriesGoal = txtCaloriesGoal.Text
        Catch ex As Exception
            CaloriesGoal = 2000
        End Try
        Try
            DistanceGoal = txtDistanceGoal.Text
        Catch ex As Exception
            DistanceGoal = 2
        End Try
        txtStepsGoal.Text = StepCountGoal
        txtCaloriesGoal.Text = CaloriesGoal
        txtDistanceGoal.Text = DistanceGoal
        Try
            Dim CurrentGoalConfig As New Goals
            Select Case CurrentBandClass
                Case BandClass.Cargo
                    CurrentGoalConfig = New Goals(chkStepsEnabled.IsChecked, chkCaloriesEnabled.IsChecked, chkDistanceEnabled.IsChecked, StepCountGoal, CaloriesGoal, DistanceGoal, Now)
                Case BandClass.Envoy
                    CurrentGoalConfig = New Goals(chkStepsEnabled.IsChecked, chkCaloriesEnabled.IsChecked, chkDistanceEnabled.IsChecked, False, StepCountGoal, CaloriesGoal, DistanceGoal, 0, Now)
                Case BandClass.Unknown
                    CurrentGoalConfig = New Goals(chkStepsEnabled.IsChecked, chkCaloriesEnabled.IsChecked, chkDistanceEnabled.IsChecked, StepCountGoal, CaloriesGoal, DistanceGoal, Now)
            End Select
            BandClient.SetGoals(CurrentGoalConfig)
        Catch ex As Exception
            MessageBox.Show("試圖設置運動目標時發生例外情況: " & ex.Message, "錯誤", MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Sub
    Private Sub chkStepsEnabled_Click(sender As Object, e As RoutedEventArgs) Handles chkStepsEnabled.Click
        txtStepsGoal.IsEnabled = chkStepsEnabled.IsChecked
    End Sub

    Private Sub chkCaloriesEnabled_Click(sender As Object, e As RoutedEventArgs) Handles chkCaloriesEnabled.Click
        txtCaloriesGoal.IsEnabled = chkCaloriesEnabled.IsChecked
    End Sub

    Private Sub chkDistanceEnabled_Click(sender As Object, e As RoutedEventArgs) Handles chkDistanceEnabled.Click
        txtDistanceGoal.IsEnabled = chkDistanceEnabled.IsChecked
    End Sub

    Private Sub txtUserFirstName_MouseUp(sender As Object, e As MouseButtonEventArgs) Handles txtUserFirstName.MouseUp
        Dim NewFirstName As String
        NewFirstName = InputBox("請輸入您的名字。", "修改名字", txtUserFirstName.Text)
        If NewFirstName.Trim() <> "" Then
            Try
                Dim NewBandUserProfile As IUserProfile
                NewBandUserProfile = BandClient.GetUserProfileFromDevice
                NewBandUserProfile.FirstName = NewFirstName
                BandClient.SaveUserProfileToBandOnly(NewBandUserProfile)
            Catch ex As Exception
                MessageBox.Show("試圖修改名字時發生例外情況: " & ex.Message, "錯誤", MessageBoxButton.OK, MessageBoxImage.Error)
            End Try
            Try
                txtUserFirstName.Text = BandClient.GetUserProfileFromDevice.FirstName
            Catch ex As Exception
                txtUserFirstName.Text = "試圖獲取資料時發生例外情況: " & ex.Message
            End Try
        End If
    End Sub
    Private Sub txtUserLastName_MouseUp(sender As Object, e As MouseButtonEventArgs) Handles txtUserLastName.MouseUp
        Dim NewLastName As String
        NewLastName = InputBox("請輸入您的姓氏。", "修改姓氏", txtUserLastName.Text)
        If NewLastName.Trim() <> "" Then
            Try
                Dim NewBandUserProfile As IUserProfile
                NewBandUserProfile = BandClient.GetUserProfileFromDevice
                NewBandUserProfile.LastName = NewLastName
                BandClient.SaveUserProfileToBandOnly(NewBandUserProfile)
            Catch ex As Exception
                MessageBox.Show("試圖修改姓氏時發生例外情況: " & ex.Message, "錯誤", MessageBoxButton.OK, MessageBoxImage.Error)
            End Try
            Try
                txtUserLastName.Text = BandClient.GetUserProfileFromDevice.LastName
            Catch ex As Exception
                txtUserLastName.Text = "試圖獲取資料時發生例外情況: " & ex.Message
            End Try
        End If
    End Sub
    Private Sub txtUserHeight_MouseUp(sender As Object, e As MouseButtonEventArgs) Handles txtUserHeight.MouseUp
        Dim NewHeight As String
        NewHeight = InputBox("請輸入您的身高。", "修改身高", txtUserHeight.Text)
        If IsNumeric(NewHeight) Then
            Try
                Dim NewBandUserProfile As IUserProfile
                NewBandUserProfile = BandClient.GetUserProfileFromDevice
                NewBandUserProfile.Height = CInt(NewHeight)
                BandClient.SaveUserProfileToBandOnly(NewBandUserProfile)
            Catch ex As Exception
                MessageBox.Show("試圖修改身高時發生例外情況: " & ex.Message, "錯誤", MessageBoxButton.OK, MessageBoxImage.Error)
            End Try
            Try
                txtUserHeight.Text = BandClient.GetUserProfileFromDevice.Height
            Catch ex As Exception
                txtUserHeight.Text = "試圖獲取資料時發生例外情況: " & ex.Message
            End Try
        End If
    End Sub
    Private Sub txtUserWeight_MouseUp(sender As Object, e As MouseButtonEventArgs) Handles txtUserWeight.MouseUp
        Dim NewWeight As String
        NewWeight = InputBox("請輸入您的體重。", "修改體重", txtUserWeight.Text)
        If IsNumeric(NewWeight) Then
            Try
                Dim NewBandUserProfile As IUserProfile
                NewBandUserProfile = BandClient.GetUserProfileFromDevice
                NewBandUserProfile.Weight = CInt(NewWeight)
                BandClient.SaveUserProfileToBandOnly(NewBandUserProfile)
            Catch ex As Exception
                MessageBox.Show("試圖修改體重時發生例外情況: " & ex.Message, "錯誤", MessageBoxButton.OK, MessageBoxImage.Error)
            End Try
            Try
                txtUserWeight.Text = BandClient.GetUserProfileFromDevice.Weight
            Catch ex As Exception
                txtUserWeight.Text = "試圖獲取資料時發生例外情況: " & ex.Message
            End Try
        End If
    End Sub
    Private Sub txtEmailAddress_MouseUp(sender As Object, e As MouseButtonEventArgs) Handles txtEmailAddress.MouseUp
        Dim NewEmailAddress As String
        NewEmailAddress = InputBox("請輸入您的電子郵件地址。", "修改電子郵件地址", txtEmailAddress.Text)
        If NewEmailAddress.Trim() <> "" Then
            Try
                Dim NewBandUserProfile As IUserProfile
                NewBandUserProfile = BandClient.GetUserProfileFromDevice
                NewBandUserProfile.EmailAddress = NewEmailAddress
                BandClient.SaveUserProfileToBandOnly(NewBandUserProfile)
            Catch ex As Exception
                MessageBox.Show("試圖修改電子郵件地址時發生例外情況: " & ex.Message, "錯誤", MessageBoxButton.OK, MessageBoxImage.Error)
            End Try
            Try
                txtEmailAddress.Text = BandClient.GetUserProfileFromDevice.EmailAddress
            Catch ex As Exception
                txtEmailAddress.Text = "試圖獲取資料時發生例外情況: " & ex.Message
            End Try
        End If
    End Sub
    Private Sub txtSmsAddress_MouseUp(sender As Object, e As MouseButtonEventArgs) Handles txtSmsAddress.MouseUp
        Dim NewSmsAddress As String
        NewSmsAddress = InputBox("請輸入您的簡訊地址。", "修改簡訊地址", txtSmsAddress.Text)
        If NewSmsAddress.Trim() <> "" Then
            Try
                Dim NewBandUserProfile As IUserProfile
                NewBandUserProfile = BandClient.GetUserProfileFromDevice
                NewBandUserProfile.SmsAddress = NewSmsAddress
                BandClient.SaveUserProfileToBandOnly(NewBandUserProfile)
            Catch ex As Exception
                MessageBox.Show("試圖修改簡訊地址時發生例外情況: " & ex.Message, "錯誤", MessageBoxButton.OK, MessageBoxImage.Error)
            End Try
            Try
                txtSmsAddress.Text = BandClient.GetUserProfileFromDevice.SmsAddress
            Catch ex As Exception
                txtSmsAddress.Text = "試圖獲取資料時發生例外情況: " & ex.Message
            End Try
        End If
    End Sub

    Private Sub rdbMale_Click(sender As Object, e As RoutedEventArgs) Handles rdbMale.Click
        Try
            Dim NewBandUserProfile As IUserProfile
            NewBandUserProfile = BandClient.GetUserProfileFromDevice
            NewBandUserProfile.Gender = Gender.Male
            BandClient.SaveUserProfileToBandOnly(NewBandUserProfile)
        Catch ex As Exception
            MessageBox.Show("試圖修改性別時發生例外情況: " & ex.Message, "錯誤", MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
        Try
            If BandClient.GetUserProfileFromDevice.Gender = Gender.Male Then
                rdbMale.IsChecked = True
                rdbFemale.IsChecked = False
            Else
                rdbMale.IsChecked = False
                rdbFemale.IsChecked = True
            End If
        Catch ex As Exception
            rdbMale.IsChecked = False
            rdbFemale.IsChecked = False
        End Try
    End Sub

    Private Sub rdbFemale_Click(sender As Object, e As RoutedEventArgs) Handles rdbFemale.Click
        Try
            Dim NewBandUserProfile As IUserProfile
            NewBandUserProfile = BandClient.GetUserProfileFromDevice
            NewBandUserProfile.Gender = Gender.Female
            BandClient.SaveUserProfileToBandOnly(NewBandUserProfile)
        Catch ex As Exception
            MessageBox.Show("試圖修改性別時發生例外情況: " & ex.Message, "錯誤", MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
        Try
            If BandClient.GetUserProfileFromDevice.Gender = Gender.Male Then
                rdbMale.IsChecked = True
                rdbFemale.IsChecked = False
            Else
                rdbMale.IsChecked = False
                rdbFemale.IsChecked = True
            End If
        Catch ex As Exception
            rdbMale.IsChecked = False
            rdbFemale.IsChecked = False
        End Try
    End Sub

    Private Sub rdbMetricDistUnit_Click(sender As Object, e As RoutedEventArgs) Handles rdbMetricDistUnit.Click
        Try
            Dim NewBandUserProfile As IUserProfile
            NewBandUserProfile = BandClient.GetUserProfileFromDevice
            NewBandUserProfile.DeviceSettings.LocaleSettings.DistanceLongUnits = DistanceUnitType.Metric
            NewBandUserProfile.DeviceSettings.LocaleSettings.DistanceShortUnits = DistanceUnitType.Metric
            BandClient.SaveUserProfileToBandOnly(NewBandUserProfile)
        Catch ex As Exception
            MessageBox.Show("試圖修改距離度量單位時發生例外情況: " & ex.Message, "錯誤", MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
        Try
            If BandClient.GetUserProfileFromDevice.DeviceSettings.LocaleSettings.DistanceLongUnits = DistanceUnitType.Metric Then
                rdbMetricDistUnit.IsChecked = True
                rdbImperialDistUnit.IsChecked = False
            Else
                rdbMetricDistUnit.IsChecked = False
                rdbImperialDistUnit.IsChecked = True
            End If
        Catch ex As Exception
            rdbMetricDistUnit.IsChecked = False
            rdbImperialDistUnit.IsChecked = False
        End Try
    End Sub

    Private Sub rdbImperialDistUnit_Click(sender As Object, e As RoutedEventArgs) Handles rdbImperialDistUnit.Click
        Try
            Dim NewBandUserProfile As IUserProfile
            NewBandUserProfile = BandClient.GetUserProfileFromDevice
            NewBandUserProfile.DeviceSettings.LocaleSettings.DistanceLongUnits = DistanceUnitType.Imperial
            NewBandUserProfile.DeviceSettings.LocaleSettings.DistanceShortUnits = DistanceUnitType.Imperial
            BandClient.SaveUserProfileToBandOnly(NewBandUserProfile)
        Catch ex As Exception
            MessageBox.Show("試圖修改距離度量單位時發生例外情況: " & ex.Message, "錯誤", MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
        Try
            If BandClient.GetUserProfileFromDevice.DeviceSettings.LocaleSettings.DistanceLongUnits = DistanceUnitType.Metric Then
                rdbMetricDistUnit.IsChecked = True
                rdbImperialDistUnit.IsChecked = False
            Else
                rdbMetricDistUnit.IsChecked = False
                rdbImperialDistUnit.IsChecked = True
            End If
        Catch ex As Exception
            rdbMetricDistUnit.IsChecked = False
            rdbImperialDistUnit.IsChecked = False
        End Try
    End Sub
    Private Sub rdbMetricMassUnit_Click(sender As Object, e As RoutedEventArgs) Handles rdbMetricMassUnit.Click
        Try
            Dim NewBandUserProfile As IUserProfile
            NewBandUserProfile = BandClient.GetUserProfileFromDevice
            NewBandUserProfile.DeviceSettings.LocaleSettings.MassUnits = MassUnitType.Metric
            BandClient.SaveUserProfileToBandOnly(NewBandUserProfile)
        Catch ex As Exception
            MessageBox.Show("試圖修改重量度量單位時發生例外情況: " & ex.Message, "錯誤", MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
        Try
            If BandClient.GetUserProfileFromDevice.DeviceSettings.LocaleSettings.MassUnits = MassUnitType.Metric Then
                rdbMetricMassUnit.IsChecked = True
                rdbImperialMassUnit.IsChecked = False
            Else
                rdbMetricMassUnit.IsChecked = False
                rdbImperialMassUnit.IsChecked = True
            End If
        Catch ex As Exception
            rdbMetricMassUnit.IsChecked = False
            rdbImperialMassUnit.IsChecked = False
        End Try
    End Sub

    Private Sub rdbImperialMassUnit_Click(sender As Object, e As RoutedEventArgs) Handles rdbImperialMassUnit.Click
        Try
            Dim NewBandUserProfile As IUserProfile
            NewBandUserProfile = BandClient.GetUserProfileFromDevice
            NewBandUserProfile.DeviceSettings.LocaleSettings.MassUnits = MassUnitType.Imperial
            BandClient.SaveUserProfileToBandOnly(NewBandUserProfile)
        Catch ex As Exception
            MessageBox.Show("試圖修改重量度量單位時發生例外情況: " & ex.Message, "錯誤", MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
        Try
            If BandClient.GetUserProfileFromDevice.DeviceSettings.LocaleSettings.MassUnits = MassUnitType.Metric Then
                rdbMetricMassUnit.IsChecked = True
                rdbImperialMassUnit.IsChecked = False
            Else
                rdbMetricMassUnit.IsChecked = False
                rdbImperialMassUnit.IsChecked = True
            End If
        Catch ex As Exception
            rdbMetricMassUnit.IsChecked = False
            rdbImperialMassUnit.IsChecked = False
        End Try
    End Sub
    Private Sub rdbMetricVolUnit_Click(sender As Object, e As RoutedEventArgs) Handles rdbMetricVolUnit.Click
        Try
            Dim NewBandUserProfile As IUserProfile
            NewBandUserProfile = BandClient.GetUserProfileFromDevice
            NewBandUserProfile.DeviceSettings.LocaleSettings.VolumeUnits = VolumeUnitType.Metric
            BandClient.SaveUserProfileToBandOnly(NewBandUserProfile)
        Catch ex As Exception
            MessageBox.Show("試圖修改體積度量單位時發生例外情況: " & ex.Message, "錯誤", MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
        Try
            If BandClient.GetUserProfileFromDevice.DeviceSettings.LocaleSettings.VolumeUnits = VolumeUnitType.Metric Then
                rdbMetricVolUnit.IsChecked = True
                rdbImperialVolUnit.IsChecked = False
            Else
                rdbMetricVolUnit.IsChecked = False
                rdbImperialVolUnit.IsChecked = True
            End If
        Catch ex As Exception
            rdbMetricVolUnit.IsChecked = False
            rdbImperialVolUnit.IsChecked = False
        End Try
    End Sub

    Private Sub rdbImperialVolUnit_Click(sender As Object, e As RoutedEventArgs) Handles rdbImperialVolUnit.Click
        Try
            Dim NewBandUserProfile As IUserProfile
            NewBandUserProfile = BandClient.GetUserProfileFromDevice
            NewBandUserProfile.DeviceSettings.LocaleSettings.VolumeUnits = VolumeUnitType.Imperial
            BandClient.SaveUserProfileToBandOnly(NewBandUserProfile)
        Catch ex As Exception
            MessageBox.Show("試圖修改體積度量單位時發生例外情況: " & ex.Message, "錯誤", MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
        Try
            If BandClient.GetUserProfileFromDevice.DeviceSettings.LocaleSettings.VolumeUnits = VolumeUnitType.Metric Then
                rdbMetricVolUnit.IsChecked = True
                rdbImperialVolUnit.IsChecked = False
            Else
                rdbMetricVolUnit.IsChecked = False
                rdbImperialVolUnit.IsChecked = True
            End If
        Catch ex As Exception
            rdbMetricVolUnit.IsChecked = False
            rdbImperialVolUnit.IsChecked = False
        End Try
    End Sub
    Private Sub rdbMetricTempUnit_Click(sender As Object, e As RoutedEventArgs) Handles rdbMetricTempUnit.Click
        Try
            Dim NewBandUserProfile As IUserProfile
            NewBandUserProfile = BandClient.GetUserProfileFromDevice
            NewBandUserProfile.DeviceSettings.LocaleSettings.TemperatureUnits = TemperatureUnitType.Metric
            BandClient.SaveUserProfileToBandOnly(NewBandUserProfile)
        Catch ex As Exception
            MessageBox.Show("試圖修改溫度度量單位時發生例外情況: " & ex.Message, "錯誤", MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
        Try
            If BandClient.GetUserProfileFromDevice.DeviceSettings.LocaleSettings.TemperatureUnits = TemperatureUnitType.Metric Then
                rdbMetricTempUnit.IsChecked = True
                rdbImperialTempUnit.IsChecked = False
            Else
                rdbMetricTempUnit.IsChecked = False
                rdbImperialTempUnit.IsChecked = True
            End If
        Catch ex As Exception
            rdbMetricTempUnit.IsChecked = False
            rdbImperialTempUnit.IsChecked = False
        End Try
    End Sub

    Private Sub rdbImperialTempUnit_Click(sender As Object, e As RoutedEventArgs) Handles rdbImperialTempUnit.Click
        Try
            Dim NewBandUserProfile As IUserProfile
            NewBandUserProfile = BandClient.GetUserProfileFromDevice
            NewBandUserProfile.DeviceSettings.LocaleSettings.TemperatureUnits = TemperatureUnitType.Imperial
            BandClient.SaveUserProfileToBandOnly(NewBandUserProfile)
        Catch ex As Exception
            MessageBox.Show("試圖修改溫度度量單位時發生例外情況: " & ex.Message, "錯誤", MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
        Try
            If BandClient.GetUserProfileFromDevice.DeviceSettings.LocaleSettings.TemperatureUnits = TemperatureUnitType.Metric Then
                rdbMetricTempUnit.IsChecked = True
                rdbImperialTempUnit.IsChecked = False
            Else
                rdbMetricTempUnit.IsChecked = False
                rdbImperialTempUnit.IsChecked = True
            End If
        Catch ex As Exception
            rdbMetricTempUnit.IsChecked = False
            rdbImperialTempUnit.IsChecked = False
        End Try
    End Sub
    Private Sub rdbMetricEnergyUnit_Click(sender As Object, e As RoutedEventArgs) Handles rdbMetricEnergyUnit.Click
        Try
            Dim NewBandUserProfile As IUserProfile
            NewBandUserProfile = BandClient.GetUserProfileFromDevice
            NewBandUserProfile.DeviceSettings.LocaleSettings.EnergyUnits = EnergyUnitType.Metric
            BandClient.SaveUserProfileToBandOnly(NewBandUserProfile)
        Catch ex As Exception
            MessageBox.Show("試圖修改能量度量單位時發生例外情況: " & ex.Message, "錯誤", MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
        Try
            If BandClient.GetUserProfileFromDevice.DeviceSettings.LocaleSettings.EnergyUnits = EnergyUnitType.Metric Then
                rdbMetricEnergyUnit.IsChecked = True
                rdbImperialEnergyUnit.IsChecked = False
            Else
                rdbMetricEnergyUnit.IsChecked = False
                rdbImperialEnergyUnit.IsChecked = True
            End If
        Catch ex As Exception
            rdbMetricEnergyUnit.IsChecked = False
            rdbImperialEnergyUnit.IsChecked = False
        End Try
    End Sub

    Private Sub rdbImperialEnergyUnit_Click(sender As Object, e As RoutedEventArgs) Handles rdbImperialEnergyUnit.Click
        Try
            Dim NewBandUserProfile As IUserProfile
            NewBandUserProfile = BandClient.GetUserProfileFromDevice
            NewBandUserProfile.DeviceSettings.LocaleSettings.EnergyUnits = EnergyUnitType.Imperial
            BandClient.SaveUserProfileToBandOnly(NewBandUserProfile)
        Catch ex As Exception
            MessageBox.Show("試圖修改能量度量單位時發生例外情況: " & ex.Message, "錯誤", MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
        Try
            If BandClient.GetUserProfileFromDevice.DeviceSettings.LocaleSettings.EnergyUnits = EnergyUnitType.Metric Then
                rdbMetricEnergyUnit.IsChecked = True
                rdbImperialEnergyUnit.IsChecked = False
            Else
                rdbMetricEnergyUnit.IsChecked = False
                rdbImperialEnergyUnit.IsChecked = True
            End If
        Catch ex As Exception
            rdbMetricEnergyUnit.IsChecked = False
            rdbImperialEnergyUnit.IsChecked = False
        End Try
    End Sub

    Private Sub btnAddTile_Click(sender As Object, e As RoutedEventArgs) Handles btnAddTile.Click
        If lstAvailableTiles.SelectedIndex > -1 Then
            UserEditedTiles.Add(BandAvailableTiles(lstAvailableTiles.SelectedIndex))
            UserEditedTilesNameList.Add(BandAvailableTilesNameList(lstAvailableTiles.SelectedIndex))
            BandAvailableTiles.RemoveAt(lstAvailableTiles.SelectedIndex)
            BandAvailableTilesNameList.RemoveAt(lstAvailableTiles.SelectedIndex)
            RefreshTileList(-2, UserEditedTilesNameList.Count - 1)
        End If
    End Sub

    Private Sub btnRemoveTile_Click(sender As Object, e As RoutedEventArgs) Handles btnRemoveTile.Click
        If lstPinnedTiles.SelectedIndex > -1 Then
            BandAvailableTiles.Add(UserEditedTiles(lstPinnedTiles.SelectedIndex))
            BandAvailableTilesNameList.Add(UserEditedTilesNameList(lstPinnedTiles.SelectedIndex))
            UserEditedTiles.RemoveAt(lstPinnedTiles.SelectedIndex)
            UserEditedTilesNameList.RemoveAt(lstPinnedTiles.SelectedIndex)
            RefreshTileList(BandAvailableTilesNameList.Count - 1, -2)
        End If
    End Sub

    Private Sub btnTileMoveUp_Click(sender As Object, e As RoutedEventArgs) Handles btnTileMoveUp.Click
        If lstPinnedTiles.SelectedIndex >= 1 Then
            Dim BandTileTemp As AdminBandTile = UserEditedTiles(lstPinnedTiles.SelectedIndex - 1)
            Dim BandTileNameTemp As String = UserEditedTilesNameList(lstPinnedTiles.SelectedIndex - 1)
            UserEditedTiles(lstPinnedTiles.SelectedIndex - 1) = UserEditedTiles(lstPinnedTiles.SelectedIndex)
            UserEditedTilesNameList(lstPinnedTiles.SelectedIndex - 1) = UserEditedTilesNameList(lstPinnedTiles.SelectedIndex)
            UserEditedTiles(lstPinnedTiles.SelectedIndex) = BandTileTemp
            UserEditedTilesNameList(lstPinnedTiles.SelectedIndex) = BandTileNameTemp
            RefreshTileList(-2, lstPinnedTiles.SelectedIndex - 1)
        End If
    End Sub

    Private Sub btnTileMoveDown_Click(sender As Object, e As RoutedEventArgs) Handles btnTileMoveDown.Click
        If lstPinnedTiles.SelectedIndex >= 0 And lstPinnedTiles.SelectedIndex < UserEditedTilesNameList.Count - 1 Then
            Dim BandTileTemp As AdminBandTile = UserEditedTiles(lstPinnedTiles.SelectedIndex + 1)
            Dim BandTileNameTemp As String = UserEditedTilesNameList(lstPinnedTiles.SelectedIndex + 1)
            UserEditedTiles(lstPinnedTiles.SelectedIndex + 1) = UserEditedTiles(lstPinnedTiles.SelectedIndex)
            UserEditedTilesNameList(lstPinnedTiles.SelectedIndex + 1) = UserEditedTilesNameList(lstPinnedTiles.SelectedIndex)
            UserEditedTiles(lstPinnedTiles.SelectedIndex) = BandTileTemp
            UserEditedTilesNameList(lstPinnedTiles.SelectedIndex) = BandTileNameTemp
            RefreshTileList(-2, lstPinnedTiles.SelectedIndex + 1)
        End If
    End Sub

    Private Sub lstAvailableTiles_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles lstAvailableTiles.SelectionChanged
        If lstAvailableTiles.SelectedIndex > -1 Then
            btnAddTile.IsEnabled = True
        End If
    End Sub

    Private Sub lstPinnedTiles_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles lstPinnedTiles.SelectionChanged
        If lstPinnedTiles.SelectedIndex > -1 Then
            btnRemoveTile.IsEnabled = True
            btnTileMoveUp.IsEnabled = True
            btnTileMoveDown.IsEnabled = True
            If lstPinnedTiles.SelectedIndex = 0 Then
                btnTileMoveUp.IsEnabled = False
            End If
            If lstPinnedTiles.SelectedIndex = UserEditedTilesNameList.Count - 1 Then
                btnTileMoveDown.IsEnabled = False
            End If
        End If
    End Sub

    Private Sub btnSetBandStartStrip_Click(sender As Object, e As RoutedEventArgs) Handles btnSetBandStartStrip.Click
        Try
            SetBandTileInfo()
        Catch ex As Exception
            MessageBox.Show("試圖設置裝置動態磚資料時發生例外情況: " & ex.Message, "錯誤", MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
        Try
            GetBandTileInfo()
            lstAvailableTiles.ItemsSource = EmptyList
            lstPinnedTiles.ItemsSource = EmptyList
            lstAvailableTiles.ItemsSource = BandAvailableTilesNameList
            lstPinnedTiles.ItemsSource = UserEditedTilesNameList
        Catch ex As Exception
            MessageBox.Show("試圖獲取裝置動態磚資料時發生例外情況: " & ex.Message, "錯誤", MessageBoxButton.OK, MessageBoxImage.Error)
            lstAvailableTiles.ItemsSource = EmptyList
            lstPinnedTiles.ItemsSource = EmptyList
        End Try
    End Sub

    Private Sub cmbDateFormat_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles cmbDateFormat.SelectionChanged
        If cmbDateFormat.SelectedIndex > 0 Then
            Try
                Dim NewBandUserProfile As IUserProfile
                NewBandUserProfile = BandClient.GetUserProfileFromDevice
                NewBandUserProfile.DeviceSettings.LocaleSettings.DateFormat = cmbDateFormat.SelectedIndex
                BandClient.SaveUserProfileToBandOnly(NewBandUserProfile)
            Catch ex As Exception
                MessageBox.Show("試圖修改日期格式時發生例外情況: " & ex.Message, "錯誤", MessageBoxButton.OK, MessageBoxImage.Error)
            End Try
            Try
                cmbDateFormat.SelectedIndex = BandClient.GetUserProfileFromDevice.DeviceSettings.LocaleSettings.DateFormat
            Catch ex As Exception
                cmbDateFormat.SelectedIndex = 0
            End Try
        End If
    End Sub

    Private Sub cmbTimeFormat_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles cmbTimeFormat.SelectionChanged
        If cmbTimeFormat.SelectedIndex > 0 Then
            Try
                Dim NewBandUserProfile As IUserProfile
                NewBandUserProfile = BandClient.GetUserProfileFromDevice
                NewBandUserProfile.DeviceSettings.LocaleSettings.TimeFormat = cmbTimeFormat.SelectedIndex
                BandClient.SaveUserProfileToBandOnly(NewBandUserProfile)
            Catch ex As Exception
                MessageBox.Show("試圖修改時間格式時發生例外情況: " & ex.Message, "錯誤", MessageBoxButton.OK, MessageBoxImage.Error)
            End Try
            Try
                cmbTimeFormat.SelectedIndex = BandClient.GetUserProfileFromDevice.DeviceSettings.LocaleSettings.TimeFormat
            Catch ex As Exception
                cmbTimeFormat.SelectedIndex = 0
            End Try
        End If
    End Sub

    Private Sub txtNumberSeparator_MouseUp(sender As Object, e As MouseButtonEventArgs) Handles txtNumberSeparator.MouseUp
        Dim NewNumberSeparator As String
        NewNumberSeparator = InputBox("請輸入新的小數符號。", "修改小數符號", txtNumberSeparator.Text)
        If NewNumberSeparator.Trim.Length >= 1 Then
            Try
                Dim NewBandUserProfile As IUserProfile
                NewBandUserProfile = BandClient.GetUserProfileFromDevice
                NewBandUserProfile.DeviceSettings.LocaleSettings.NumberSeparator = NewNumberSeparator(0)
                BandClient.SaveUserProfileToBandOnly(NewBandUserProfile)
            Catch ex As Exception
                MessageBox.Show("試圖修改小數符號時發生例外情況: " & ex.Message, "錯誤", MessageBoxButton.OK, MessageBoxImage.Error)
            End Try
            Try
                txtNumberSeparator.Text = BandClient.GetUserProfileFromDevice.DeviceSettings.LocaleSettings.NumberSeparator
            Catch ex As Exception
                txtNumberSeparator.Text = "試圖獲取資料時發生例外情況: " & ex.Message
            End Try
        End If
    End Sub

    Private Sub txtDecimalSeparator_MouseUp(sender As Object, e As MouseButtonEventArgs) Handles txtDecimalSeparator.MouseUp
        Dim NewDecimalSeparator As String
        NewDecimalSeparator = InputBox("請輸入新的數字分位符號。", "修改數字分位符號", txtDecimalSeparator.Text)
        If NewDecimalSeparator.Trim.Length >= 1 Then
            Try
                Dim NewBandUserProfile As IUserProfile
                NewBandUserProfile = BandClient.GetUserProfileFromDevice
                NewBandUserProfile.DeviceSettings.LocaleSettings.DecimalSeparator = NewDecimalSeparator(0)
                BandClient.SaveUserProfileToBandOnly(NewBandUserProfile)
            Catch ex As Exception
                MessageBox.Show("試圖修改數字分位符號時發生例外情況: " & ex.Message, "錯誤", MessageBoxButton.OK, MessageBoxImage.Error)
            End Try
            Try
                txtDecimalSeparator.Text = BandClient.GetUserProfileFromDevice.DeviceSettings.LocaleSettings.DecimalSeparator
            Catch ex As Exception
                txtDecimalSeparator.Text = "試圖獲取資料時發生例外情況: " & ex.Message
            End Try
        End If
    End Sub
    Private Sub txtDateSeparator_MouseUp(sender As Object, e As MouseButtonEventArgs) Handles txtDateSeparator.MouseUp
        Dim NewDateSeparator As String
        NewDateSeparator = InputBox("請輸入新的日期分位符號。", "修改日期分位符號", txtDateSeparator.Text)
        If NewDateSeparator.Trim.Length >= 1 Then
            Try
                Dim NewBandUserProfile As IUserProfile
                NewBandUserProfile = BandClient.GetUserProfileFromDevice
                NewBandUserProfile.DeviceSettings.LocaleSettings.DateSeparator = NewDateSeparator(0)
                BandClient.SaveUserProfileToBandOnly(NewBandUserProfile)
            Catch ex As Exception
                MessageBox.Show("試圖修改日期分位符號時發生例外情況: " & ex.Message, "錯誤", MessageBoxButton.OK, MessageBoxImage.Error)
            End Try
            Try
                txtDateSeparator.Text = BandClient.GetUserProfileFromDevice.DeviceSettings.LocaleSettings.DateSeparator
            Catch ex As Exception
                txtDateSeparator.Text = "試圖獲取資料時發生例外情況: " & ex.Message
            End Try
        End If
    End Sub

    Private Async Sub btnVibrate_Click(sender As Object, e As RoutedEventArgs) Handles btnVibrate.Click
        Try
            Await BandClient.VibrateAsync(AdminVibrationType.SystemStartUp)
        Catch ex As Exception

        End Try
    End Sub

    Private Sub btnLoadBandTheme_Click(sender As Object, e As RoutedEventArgs) Handles btnLoadBandTheme.Click
        Dim OpenThemeFileDialog As New CommonOpenFileDialog
        With OpenThemeFileDialog
            .Title = "載入佈景主題"
            .EnsureFileExists = True
            .Filters.Add(New CommonFileDialogFilter("裝置佈景主題文件", "msbandtheme"))
            .Filters.Add(New CommonFileDialogFilter("XML 文件", "xml"))
            .Filters.Add(New CommonFileDialogFilter("所有檔案", "*"))
        End With
        If OpenThemeFileDialog.ShowDialog() = CommonFileDialogResult.Ok Then
            Dim NewTheme As New BandThemeData
            Try
                NewTheme = LoadBandThemeFromFile(OpenThemeFileDialog.FileName)
                btnColorBase.Background = New SolidColorBrush(Color.FromRgb(NewTheme.BaseColor.Red, NewTheme.BaseColor.Green, NewTheme.BaseColor.Blue))
                btnColorHighContrast.Background = New SolidColorBrush(Color.FromRgb(NewTheme.HighContrastColor.Red, NewTheme.HighContrastColor.Green, NewTheme.HighContrastColor.Blue))
                btnColorHighlight.Background = New SolidColorBrush(Color.FromRgb(NewTheme.HighlightColor.Red, NewTheme.HighlightColor.Green, NewTheme.HighlightColor.Blue))
                btnColorLowlight.Background = New SolidColorBrush(Color.FromRgb(NewTheme.LowlightColor.Red, NewTheme.LowlightColor.Green, NewTheme.LowlightColor.Blue))
                btnColorMuted.Background = New SolidColorBrush(Color.FromRgb(NewTheme.MutedColor.Red, NewTheme.MutedColor.Green, NewTheme.MutedColor.Blue))
                btnColorSecondary.Background = New SolidColorBrush(Color.FromRgb(NewTheme.SecondaryColor.Red, NewTheme.SecondaryColor.Green, NewTheme.SecondaryColor.Blue))
                btnColorBase.Tag = Color.FromRgb(NewTheme.BaseColor.Red, NewTheme.BaseColor.Green, NewTheme.BaseColor.Blue)
                btnColorHighContrast.Tag = Color.FromRgb(NewTheme.HighContrastColor.Red, NewTheme.HighContrastColor.Green, NewTheme.HighContrastColor.Blue)
                btnColorHighlight.Tag = Color.FromRgb(NewTheme.HighlightColor.Red, NewTheme.HighlightColor.Green, NewTheme.HighlightColor.Blue)
                btnColorLowlight.Tag = Color.FromRgb(NewTheme.LowlightColor.Red, NewTheme.LowlightColor.Green, NewTheme.LowlightColor.Blue)
                btnColorMuted.Tag = Color.FromRgb(NewTheme.MutedColor.Red, NewTheme.MutedColor.Green, NewTheme.MutedColor.Blue)
                btnColorSecondary.Tag = Color.FromRgb(NewTheme.SecondaryColor.Red, NewTheme.SecondaryColor.Green, NewTheme.SecondaryColor.Blue)
            Catch ex As Exception
                MessageBox.Show("打開佈景主題""" & OpenThemeFileDialog.FileName & """時發生例外情況: " & ex.Message, "錯誤", MessageBoxButton.OK, MessageBoxImage.Error)
            End Try
        End If
    End Sub

    Private Sub btnSaveBandTheme_Click(sender As Object, e As RoutedEventArgs) Handles btnSaveBandTheme.Click
        Dim SaveThemeFileDialog As New CommonSaveFileDialog
        With SaveThemeFileDialog
            .Title = "儲存佈景主題"
            .EnsurePathExists = True
            .EnsureValidNames = True
            .Filters.Add(New CommonFileDialogFilter("裝置佈景主題文件", "msbandtheme"))
            .Filters.Add(New CommonFileDialogFilter("XML 文件", "xml"))
            .Filters.Add(New CommonFileDialogFilter("所有檔案", "*"))
        End With
        If SaveThemeFileDialog.ShowDialog = CommonFileDialogResult.Ok Then
            Dim ThemeToSave As New BandThemeData
            With ThemeToSave
                Dim CurrentProcessingColor As Color
                CurrentProcessingColor = btnColorBase.Tag
                .BaseColor.Red = CurrentProcessingColor.R
                .BaseColor.Green = CurrentProcessingColor.G
                .BaseColor.Blue = CurrentProcessingColor.B
                CurrentProcessingColor = btnColorHighContrast.Tag
                .HighContrastColor.Red = CurrentProcessingColor.R
                .HighContrastColor.Green = CurrentProcessingColor.G
                .HighContrastColor.Blue = CurrentProcessingColor.B
                CurrentProcessingColor = btnColorHighlight.Tag
                .HighlightColor.Red = CurrentProcessingColor.R
                .HighlightColor.Green = CurrentProcessingColor.G
                .HighlightColor.Blue = CurrentProcessingColor.B
                CurrentProcessingColor = btnColorLowlight.Tag
                .LowlightColor.Red = CurrentProcessingColor.R
                .LowlightColor.Green = CurrentProcessingColor.G
                .LowlightColor.Blue = CurrentProcessingColor.B
                CurrentProcessingColor = btnColorMuted.Tag
                .MutedColor.Red = CurrentProcessingColor.R
                .MutedColor.Green = CurrentProcessingColor.G
                .MutedColor.Blue = CurrentProcessingColor.B
                CurrentProcessingColor = btnColorSecondary.Tag
                .SecondaryColor.Red = CurrentProcessingColor.R
                .SecondaryColor.Green = CurrentProcessingColor.G
                .SecondaryColor.Blue = CurrentProcessingColor.B
            End With
            Try
                Select Case SaveThemeFileDialog.SelectedFileTypeIndex
                    Case 1
                        SaveBandThemeToFile(ThemeToSave, SaveThemeFileDialog.FileName & ".msbandtheme")
                        MessageBox.Show("成功將目前的佈景主題儲存到""" & SaveThemeFileDialog.FileName & ".msbandtheme" & """。", "儲存佈景主題", MessageBoxButton.OK, MessageBoxImage.Information)
                    Case 2
                        SaveBandThemeToFile(ThemeToSave, SaveThemeFileDialog.FileName & ".xml")
                        MessageBox.Show("成功將目前的佈景主題儲存到""" & SaveThemeFileDialog.FileName & ".xml" & """。", "儲存佈景主題", MessageBoxButton.OK, MessageBoxImage.Information)
                    Case Else
                        SaveBandThemeToFile(ThemeToSave, SaveThemeFileDialog.FileName)
                        MessageBox.Show("成功將目前的佈景主題儲存到""" & SaveThemeFileDialog.FileName & """。", "儲存佈景主題", MessageBoxButton.OK, MessageBoxImage.Information)
                End Select
            Catch ex As Exception
                MessageBox.Show("儲存目前的佈景主題到""" & SaveThemeFileDialog.FileName & """時發生例外情況: " & ex.Message, "錯誤", MessageBoxButton.OK, MessageBoxImage.Error)
            End Try
        End If
    End Sub
End Class
