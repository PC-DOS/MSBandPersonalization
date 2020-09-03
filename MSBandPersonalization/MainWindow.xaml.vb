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
    Dim BandClient As ICargoClient
    Dim EmptyList As New List(Of String)
    Dim IsDeviceConnected As Boolean = False
    Dim BandTileList As BandTile()
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
        txtUserFirstName.Text = ""
        txtUserLastName.Text = ""
        txtUserHeight.Text = ""
        txtUserWeight.Text = ""
        txtEmailAddress.Text = ""
        txtSmsAddress.Text = ""
        rdbMale.IsChecked = False
        rdbFemale.IsChecked = False
        rdbMetricUnit.IsChecked = False
        rdbAutoUnit.IsChecked = False
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
        txtUserFirstName.IsEnabled = False
        txtUserLastName.IsEnabled = False
        txtUserHeight.IsEnabled = False
        txtUserWeight.IsEnabled = False
        txtEmailAddress.IsEnabled = False
        txtSmsAddress.IsEnabled = False
        rdbMale.IsEnabled = False
        rdbFemale.IsEnabled = False
        rdbMetricUnit.IsEnabled = False
        rdbAutoUnit.IsEnabled = False
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
    End Sub
    Private Sub UnockOperationWindow()
        btnFinalizeOOBE.IsEnabled = True
        btnFactoryReset.IsEnabled = True
        btnEnableRetailDemoMode.IsEnabled = True
        btnDisableRetailDemoMode.IsEnabled = True
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
        rdbMetricUnit.IsEnabled = True
        rdbAutoUnit.IsEnabled = True
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
                If BandClient.GetUserProfileFromDevice.DeviceSettings.RunDisplayUnits = RunMeasurementUnitType.Metric Then
                    rdbMetricUnit.IsChecked = True
                    rdbAutoUnit.IsChecked = False
                Else
                    rdbMetricUnit.IsChecked = False
                    rdbAutoUnit.IsChecked = True
                End If
            Catch ex As Exception
                rdbMetricUnit.IsChecked = False
                rdbAutoUnit.IsChecked = False
            End Try
            'BandTileNameList.Clear()
            'Try
            '    BandTileList = Await BandClient.TileManager.GetTilesAsync()
            '    For Each SingleTile As BandTile In BandTileList
            '        BandTileNameList.Add(SingleTile.Name & " [GUID=" & SingleTile.TileId.ToString() & "]")
            '    Next
            'Catch ex As Exception
            '    MessageBox.Show("試圖獲取裝置動態磚資料時發生例外情況: " & ex.Message, "錯誤", MessageBoxButton.OK, MessageBoxImage.Error)
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
            Dim CurrentGoalConfig As New Goals(chkStepsEnabled.IsChecked, chkCaloriesEnabled.IsChecked, chkDistanceEnabled.IsChecked, StepCountGoal, CaloriesGoal, DistanceGoal, Now)
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

    Private Sub rdbMetricUnit_Click(sender As Object, e As RoutedEventArgs) Handles rdbMetricUnit.Click
        Try
            Dim NewBandUserProfile As IUserProfile
            NewBandUserProfile = BandClient.GetUserProfileFromDevice
            NewBandUserProfile.DeviceSettings.RunDisplayUnits = RunMeasurementUnitType.Metric
            BandClient.SaveUserProfileToBandOnly(NewBandUserProfile)
        Catch ex As Exception
            MessageBox.Show("試圖修改度量單位時發生例外情況: " & ex.Message, "錯誤", MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
        Try
            If BandClient.GetUserProfileFromDevice.DeviceSettings.RunDisplayUnits = RunMeasurementUnitType.Metric Then
                rdbMetricUnit.IsChecked = True
                rdbAutoUnit.IsChecked = False
            Else
                rdbMetricUnit.IsChecked = False
                rdbAutoUnit.IsChecked = True
            End If
        Catch ex As Exception
            rdbMetricUnit.IsChecked = False
            rdbAutoUnit.IsChecked = False
        End Try
    End Sub

    Private Sub rdbAutoUnit_Click(sender As Object, e As RoutedEventArgs) Handles rdbAutoUnit.Click
        Try
            Dim NewBandUserProfile As IUserProfile
            NewBandUserProfile = BandClient.GetUserProfileFromDevice
            NewBandUserProfile.DeviceSettings.RunDisplayUnits = RunMeasurementUnitType.UseLocaleSetting
            BandClient.SaveUserProfileToBandOnly(NewBandUserProfile)
        Catch ex As Exception
            MessageBox.Show("試圖修改度量單位時發生例外情況: " & ex.Message, "錯誤", MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
        Try
            If BandClient.GetUserProfileFromDevice.DeviceSettings.RunDisplayUnits = RunMeasurementUnitType.Metric Then
                rdbMetricUnit.IsChecked = True
                rdbAutoUnit.IsChecked = False
            Else
                rdbMetricUnit.IsChecked = False
                rdbAutoUnit.IsChecked = True
            End If
        Catch ex As Exception
            rdbMetricUnit.IsChecked = False
            rdbAutoUnit.IsChecked = False
        End Try
    End Sub
End Class
