   M S B A M L   `   `   `  ���� H  DMahApps.Metro, Version=1.6.5.1, Culture=neutral, PublicKeyToken=nullV;clr-namespace:MahApps.Metro.Controls;assembly=MahApps.MetroMahApps.Metro.Controls   MSBandPersonalization&    MSBandPersonalization.MainWindow   \ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     PControls;clr-namespace:MahApps.Metro.Controls;assembly=MahApps.Metro       � Title$;  5Microsoft Band Personalization Tool - PC-DOS Workshop��$	��500��$	��750�� � WindowStartupLocation$ CenterScreen=�$	��750��$	��500�� � Icon$ Assets/Icon.ico��.��� $��10q� � ColumnDefinitions �� $��* �$	��150��$	��245���� $
��Auto ��� $
��2.5* �.+�� $��0��$��0,0,10,0q� � RowDefinitions �� $
��Auto ��� $��* ��� $��30 �.+��� $��0��$��搜尋到的裝置: ���� -    ��Name$ 
lstDevices��$��1��$
��0,10q�#��� ��$��1q�� $��2�� �� $��* ��� $��* �.+��� -    ��Name$ btnRefreshDevice��$��0��$��重新整理(_R)��$��0,0,5,0q��� -   $ 
btnConnect��$��1��$��
連線(_C)��$��5,0,0,0q����	#FFCCCCCC�� $��1��$��Stretch=�$��Center=�$	��2.5��$��0q�,   &MahApps.Metro.Controls.MetroTabControl  $��2��.@�)   #MahApps.Metro.Controls.MetroTabItem  $��裝置資訊��+   %MahApps.Metro.Controls.ControlsHelper   HeaderFontSize$ 20Q�.��� $��10,0,0,0q�.+��� .����  �� Orientation$ Vertical=�.+�� $��0,0,0,5q� �� $��1.25* ��� $
��2.5* �.+��� $��0��$��裝置型號:��$
��Bold"�$
��Wrap=��� -   	 ��Name$	 txtBandClass��$��1��$�� ��$
��Wrap=�� $��0,0,0,5q� �� $��1.25* ��� $
��2.5* �.+��� $��0��$��裝置序號:��$
��Bold"�$
��Wrap=��� -   $	 txtSerialNumber��$��1��$�� ��$
��Wrap=�� $��0,0,0,5q� �� $��1.25* ��� $
��2.5* �.+��� $��0��$��裝置 GUID:��$
��Bold"�$
��Wrap=��� -   $	 txtDeviceUniqueID��$��1��$�� ��$
��Wrap=�� $��0,0,0,5q� �� $��1.25* ��� $
��2.5* �.+��� $��0��$��硬體修訂編號:��$
��Bold"�$
��Wrap=��� -   $	 txtHardwareVersion��$��1��$�� ��$
��Wrap=�� $��0,0,0,5q� �� $��1.25* ��� $
��2.5* �.+��� $��0��$��開機載入器版本:��$
��Bold"�$
��Wrap=��� -   $	 txtBootloaderVersion��$��1��$�� ��$
��Wrap=�� $��0,0,0,5q� �� $��1.25* ��� $
��2.5* �.+��� $��0��$��應用程式版本:��$
��Bold"�$
��Wrap=��� -	   $	 txtApplicationVersion��$��1��$�� ��$
��Wrap=�� $��0,0,0,5q� �� $��1.25* ��� $
��2.5* �.+��� $��0��$��更新程式版本:��$
��Bold"�$
��Wrap=��� -
   $	 txtUpdaterVersion��$��1��$�� ��$
��Wrap=�� $��0,0,0,5q� �� $��1.25* ��� $
��2.5* �.+��� $��0��$��裝置 PCBID:��$
��Bold"�$
��Wrap=��� -   $	 txtPCBID��$��1��$�� ��$
��Wrap=�� $��0,0,0,5q� �� $��1.25* ��� $
��2.5* �.+��� $��0��$��OOBE 狀態:��$
��Bold"�$
��Wrap=��� -   $	 txtOOBEStage��$��1��$�� ��$
��Wrap=�� $��0,0,0,5q� �� $��1.25* ��� $
��2.5* �.+��� $��0��$��OOBE 已完成:��$
��Bold"�$
��Wrap=��� -   $	 txtIsOOBECompleted��$��1��$�� ��$
��Wrap=��� -   
 ��Name$
 stkCargoOnly��$ Vertical=�.+�� $��0,0,0,5q� �� $��1.25* ��� $
��2.5* �.+��� $��0��$+��%裝置名稱 (點選以重新命名):��$
��Bold"�$
��Wrap=��� -   $	 txtDeviceName��$��1��$�� ��$
��Wrap=�� $��0,0,0,5q� �� $��1.25* ��� $
��2.5* �.+��� $��0��$%��使用者名 (點選以修改):��$
��Bold"�$
��Wrap=��� -   $	 txtUserFirstName��$��1��$�� ��$
��Wrap=�� $��0,0,0,5q� �� $��1.25* ��� $
��2.5* �.+��� $��0��$%��使用者姓 (點選以修改):��$
��Bold"�$
��Wrap=��� -   $	 txtUserLastName��$��1��$�� ��$
��Wrap=�� $��0,0,0,5q� �� $��1.25* ��� $
��2.5* �.+��� $��0��$��身高 (點選以修改):��$
��Bold"�$
��Wrap=��� -   $	 txtUserHeight��$��1��$�� ��$
��Wrap=�� $��0,0,0,5q� �� $��1.25* ��� $
��2.5* �.+��� $��0��$��體重 (點選以修改):��$
��Bold"�$
��Wrap=��� -   $	 txtUserWeight��$��1��$�� ��$
��Wrap=�� $��0,0,0,5q� �� $��1.25* ��� $
��2.5* �.+��� $��0��$+��%電子郵件地址 (點選以修改):��$
��Bold"�$
��Wrap=��� -   $	 txtEmailAddress��$��1��$�� ��$
��Wrap=�� $��0,0,0,5q� �� $��1.25* ��� $
��2.5* �.+��� $��0��$%��簡訊地址 (點選以修改):��$
��Bold"�$
��Wrap=��� -   $	 txtSmsAddress��$��1��$�� ��$
��Wrap=�� $��0,0,0,5q� �� $��1.25* ��� $
��2.5* �.+��� $��0��$��性別:��$
��Bold"�$
��Wrap=�� $��1�� �� $��* ��� $��* �.+�	� -    	�Name$ rdbMale��$��0��$��男性��	� -   $ 	rdbFemale��$��1��$��女性��� $��0,0,0,5q� �� $��1.25* ��� $
��2.5* �.+��� $��0��$��距離度量單位:��$
��Bold"�$
��Wrap=�� $��1�� �� $��* ��� $��* �.+�	� -   $ rdbMetricDistUnit��$��0��$��公制��	� -   $ rdbImperialDistUnit��$��1��$��英制��� $��0,0,0,5q� �� $��1.25* ��� $
��2.5* �.+��� $��0��$��重量度量單位:��$
��Bold"�$
��Wrap=�� $��1�� �� $��* ��� $��* �.+�	� -   $ rdbMetricMassUnit��$��0��$��公制��	� -   $ rdbImperialMassUnit��$��1��$��英制��� $��0,0,0,5q� �� $��1.25* ��� $
��2.5* �.+��� $��0��$��體積度量單位:��$
��Bold"�$
��Wrap=�� $��1�� �� $��* ��� $��* �.+�	� -   $ rdbMetricVolUnit��$��0��$��公制��	� -   $ rdbImperialVolUnit��$��1��$��英制��� $��0,0,0,5q� �� $��1.25* ��� $
��2.5* �.+��� $��0��$��溫度度量單位:��$
��Bold"�$
��Wrap=�� $��1�� �� $��* ��� $��* �.+�	� -   $ rdbMetricTempUnit��$��0��$��公制��	� -   $ rdbImperialTempUnit��$��1��$��英制��� $��0,0,0,5q� �� $��1.25* ��� $
��2.5* �.+��� $��0��$��能量度量單位:��$
��Bold"�$
��Wrap=�� $��1�� �� $��* ��� $��* �.+�	� -    $ rdbMetricEnergyUnit��$��0��$��公制��	� -!   $ rdbImperialEnergyUnit��$��1��$��英制��� $��0,0,0,5q� �� $��1.25* ��� $
��2.5* �.+��� $��0��$��日期格式:��$
��Bold"�$
��Wrap=��� -"    ��Name$ cmbDateFormat��$��1��.@��� $��	未指定���� $��yyyyMMdd���� $��ddMMyyyy���� $��dMMyyyy���� $��MMddyyyy���� $��Mdyyyy��� $��0,0,0,5q� �� $��1.25* ��� $
��2.5* �.+��� $��0��$��時間格式:��$
��Bold"�$
��Wrap=��� -#   $ cmbTimeFormat��$��1��.@��� $��	未指定���� $��HHmmss���� $��Hmmss��� $��0,0,0,5q� �� $��1.25* ��� $
��2.5* �.+��� $��0��$%��小數符號 (點選以修改):��$
��Bold"�$
��Wrap=��� -$   $	 txtDecimalSeparator��$��1��$�� ��$
��Wrap=�� $��0,0,0,5q� �� $��1.25* ��� $
��2.5* �.+��� $��0��$+��%數字分位符號 (點選以修改):��$
��Bold"�$
��Wrap=��� -%   $	 txtNumberSeparator��$��1��$�� ��$
��Wrap=�� $��0,0,0,5q� �� $��1.25* ��� $
��2.5* �.+��� $��0��$+��%日期分位符號 (點選以修改):��$
��Bold"�$
��Wrap=��� -&   $	 txtDateSeparator��$��1��$�� ��$
��Wrap=��� -'   $ btnFinalizeOOBE��$��強制結束 OOBE��$	��100��$
��Left=�$��0,0,0,10q��� -(   $ btnFactoryReset��$��恢復原廠設定��$��0,0,0,10q�$	��100��$
��Left=��� -)   $ btnEnableRetailDemoMode��$��啟用示範模式��$��0,0,0,10q�$	��100��$
��Left=��� -*   $ btnDisableRetailDemoMode��$��停用示範模式��$��0,0,0,10q�$	��100��$
��Left=��� -+   $ 
btnVibrate��$��讓裝置震動��$��0,0,0,10q�$	��100��$
��Left=�  $	��磚��$ 20Q�.��� $��10,0,0,0q� �� $��* ��� $
��Auto �.+�� $��0�� �� $��* ��� $
��Auto ��� $��* ��� $
��Auto �.+�� $��0�� �� $
��Auto ��� $��* �.+��� $��0��$��可用的磚: ���� -,   $ lstAvailableTiles��$��1��#��� ��$��1q�$��0,5,0,0q��� $ Vertical=�$��1��$��Center=�$	��5,0q�.+��� --   $ 
btnAddTile��$��	添加 ->��$��30���� -.   $ btnRemoveTile��$��	<- 刪除��$��0,10,0,0q�$��30��� $��2�� �� $
��Auto ��� $��* �.+��� $��0��$��已釘選的磚: ���� -/   $ lstPinnedTiles��$��1��#��� ��$��1q�$��0,5,0,0q��� $ Vertical=�$��3��$��Center=�$��5,0,0,0q�.+��� -0   $ btnTileMoveUp��$��
↑ 上移��$��30���� -1   $ btnTileMoveDown��$��
↓ 下移��$��30��$��0,10,0,0q��� -2   $ btnSetBandStartStrip��$��1��$��設置動態磚布局(_L)��$��0,10,0,0q�$��Right=�$��30��  $��運動目標��$ 20Q�.��� $��10,0,0,0q�.+��� .���� $ Vertical=�.+�� $��0,0,0,5q� �� $��1.25* ��� $
��2.5* �.+��� -3    ��Name$ chkStepsEnabled��$��設置步數目標(_P)��$��0���� -4    ��Name$ txtStepsGoal��$��1��$
��5000��{��� *	IsChecked+ �� ElementName$ chkStepsEnabled��� $��0,0,0,5q� �� $��1.25* ��� $
��2.5* �.+��� -5   $ chkCaloriesEnabled��$��設置卡路里目標(_L)��$��0���� -6   $ txtCaloriesGoal��$��1��$
��2000��{��� *	IsChecked+$ chkCaloriesEnabled��� $��0,0,0,5q� �� $��1.25* ��� $
��2.5* �.+��� -7   $ chkDistanceEnabled��$��設置路程目標(_D)��$��0���� -8   $ txtDistanceGoal��$��1��$��2��{��� *	IsChecked+$ chkDistanceEnabled���� -9   $ 
btnSetGoal��$%��儲存運動目標到裝置(_G)��$��Right=�$��0,10,0,0q�$��30��$	��125��  $��主動態磚背景圖片��$ 20Q�.��� $��10,0,0,0q� �� $
��Auto ��� $
��Auto ��� $
��Auto �.+��� $��0��$ Vertical=�.+��� #��� ��$
��Wrap=�$��0,0,0,10q��� $f�`開始前，請確保您已經將主動態磚背景圖片變更為您的裝置支援的尺寸。���� �� $���如果您使用的是第一代 Microsoft Band，您應當使用尺寸為 310 x 102，即寬度為 310 個像素、高度為 102 個像素的圖片。���� �� $���如果您使用的是第二代 Microsoft Band，即 Microsoft Band 2，您應當使用尺寸為 310 x 128，即寬度為 310 個像素、高度為 128 個像素的圖片。���� $(��"目前的主動態磚背景圖片:��#��� ���� -:    ��Name$ imgMeTileImage��$��1��$
��0,10q�$
��None=�$
��Left=��� $��2��$ 
Horizontal=�$��30��.+��� -;   $ btnBrowseMeTileImage��$��瀏覽(_W)...���� -<   $ btnSetMeTileImage��$(��"設置主動態磚背景圖片(_S)��$��10,0,0,0q�  $��佈景主題��$ 20Q�.��� $��10,0,0,0q� �� $��* ��� $
��Auto �.+��� $��0��.���� $ Vertical=�.+�� $��0,0,0,5q� �� $��1.25* ��� $
��1.5* �.+��� $��0��$��基礎顏色 (Base):��$
��Bold"�$
��Wrap=�$��Center=�� -=    �Name$ btnColorBase��$��1�����	#FF00AFF5$��25��� $��0,0,0,5q� �� $��1.25* ��� $
��1.5* �.+��� $��0��$��強調顏色 (Highlight):��$
��Bold"�$
��Wrap=�$��Center=�� ->   $ btnColorHighlight��$��1�����	#FF00AFF5$��25��� $��0,0,0,5q� �� $��1.25* ��� $
��1.5* �.+��� $��0��$��弱化顏色 (Lowlight):��$
��Bold"�$
��Wrap=�$��Center=�� -?   $ btnColorLowlight��$��1�����	#FF0096F5$��25��� $��0,0,0,5q� �� $��1.25* ��� $
��1.5* �.+��� $��0��$��第二顏色 (Secondary):��$
��Bold"�$
��Wrap=�$��Center=�� -@   $ btnColorSecondary��$��1�����	#FFE1E1E1$��25��� $��0,0,0,5q� �� $��1.25* ��� $
��1.5* �.+��� $��0��$(��"高對比度顏色 (HighContrast):��$
��Bold"�$
��Wrap=�$��Center=�� -A   $ btnColorHighContrast��$��1�����	#FF00C8F5$��25��� $��0,0,0,5q� �� $��1.25* ��� $
��1.5* �.+��� $��0��$��柔和顏色 (Muted):��$
��Bold"�$
��Wrap=�$��Center=�� -B   $ btnColorMuted��$��1�����	#FF0096C8$��25���� $��Assets/UIAssets/Color.png��$��0,10,0,0q��� -C   $ btnSetBandTheme��$��1��$"��設置裝置佈景主題(_T)��$��Right=�$��30��$	��125��$��0,10,0,0q�