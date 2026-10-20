Public Class PageSetupSystem


    Private Sub PageSetupSystem_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded

        '重复加载部分
        PanBack.ScrollToHome()
        RefreshRegistryIsolationCombo()

        If BuildType = BuildTypes.Release Then
            PanDonate.Visibility = Visibility.Collapsed
        Else
            PanDonate.Visibility = Visibility.Visible
            ItemSystemUpdateDownload.Content = "在有新版本时自动下载（更新快照版前需要先输入土豆码）"
        End If

        '非重复加载部分
        Static IsLoaded As Boolean = False
        If IsLoaded Then Return
        IsLoaded = True

        AniControlEnabled += 1
        Reload()
        SliderLoad()
        AniControlEnabled -= 1

    End Sub
    Public Sub Reload()
        SettingService.RefreshSettings(Me)
    End Sub
    Public Sub Reset()
        Try
            SettingService.ResetSettings(Me)
            Logger.Info("已初始化启动器页设置")
            Hint("已初始化启动器页设置！", HintType.Green, False)
        Catch ex As Exception
            Logger.Error(ex, "初始化启动器页设置失败", LogBehavior.Alert)
        End Try
        Reload()
    End Sub

    '注册表隔离（更改设置在下次启动生效，当前进程不切换设置路径）
    Private _refreshingRegistryIsolation As Boolean = False
    Private Sub RefreshRegistryIsolationCombo()
        _refreshingRegistryIsolation = True
        Try
            Dim Saved = GetRegistryIsolationChoice()
            ComboSystemRegistryIsolation.SelectedIndex = If(Saved.GetValueOrDefault(True), 0, 1)
            ComboSystemRegistryIsolation.ToolTip = If(Saved.HasValue AndAlso Saved.Value <> RegistryIsolated,
                "设置已保存，将在下次启动 PCL 时生效。", "修改后重新启动 PCL 生效")
        Finally
            _refreshingRegistryIsolation = False
        End Try
    End Sub

    Private Sub ComboSystemRegistryIsolation_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles ComboSystemRegistryIsolation.SelectionChanged
        If _refreshingRegistryIsolation OrElse Not ComboSystemRegistryIsolation.IsLoaded Then Return
        Dim Isolate = ComboSystemRegistryIsolation.SelectedIndex = 0
        Dim Saved = GetRegistryIsolationChoice()
        If Saved.HasValue AndAlso Saved.Value = Isolate Then Return
        Dim Answer = System.Windows.MessageBox.Show(
            If(Isolate, "是否启用注册表隔离，改用独立设置？", "是否关闭注册表隔离，与 PCL 正式版共用注册表设置？") &
            vbCrLf & vbCrLf & "更改将在重启后生效，不会自动迁移或删除设置文件。",
            "确认修改注册表隔离", System.Windows.MessageBoxButton.YesNo, System.Windows.MessageBoxImage.Question)
        If Answer <> System.Windows.MessageBoxResult.Yes OrElse Not TrySaveRegistryIsolationChoice(Isolate) Then
            RefreshRegistryIsolationCombo()
            Return
        End If
        RefreshRegistryIsolationCombo()
        MyMsgBox("注册表隔离选项已保存。重新启动 PCL 后生效。", "设置已保存")
    End Sub

    '滑动条
    Private Sub SliderLoad()
        SliderDownloadThread.GetHintText = Function(v) v + 1
        SliderDownloadSpeed.GetHintText =
        Function(v)
            Select Case v
                Case Is <= 14
                    Return (v + 1) * 0.1 & " M/s"
                Case Is <= 31
                    Return (v - 11) * 0.5 & " M/s"
                Case Is <= 41
                    Return (v - 21) & " M/s"
                Case Else
                    Return "无限制"
            End Select
        End Function
        SliderDebugAnim.GetHintText = Function(v) If(v > 29, "关闭", (v / 10 + 0.1) & "x")
    End Sub
    Private Sub SliderDownloadThread_PreviewChange(sender As Object, e As RouteEventArgs) Handles SliderDownloadThread.PreviewChange
        If SliderDownloadThread.Value < 100 Then Return
        If Not Settings.Get(Of Boolean)("HintDownloadThread") Then
            Settings.Set("HintDownloadThread", True)
            MyMsgBox("如果设置过多的下载线程，可能会导致下载时出现非常严重的卡顿。" & vbCrLf &
                     "一般设置 64 线程即可满足大多数下载需求，除非你知道你在干什么，否则不建议设置更多的线程数！", "警告", "我知道了", IsWarn:=True)
        End If
    End Sub

    '识别码/土豆码替代入口
    Private Sub BtnSystemIdentify_Click(sender As Object, e As MouseButtonEventArgs) Handles BtnSystemIdentify.Click
        PageOtherAbout.CopyIdentify()
    End Sub
    Private Sub BtnSystemUnlock_Click(sender As Object, e As MouseButtonEventArgs) Handles BtnSystemUnlock.Click
        InputPotatoCode(False)
    End Sub

    '调试模式
    Private Sub CheckDebugMode_Change() Handles CheckDebugMode.Change
        If AniControlEnabled = 0 Then Hint("部分调试信息将在刷新或启动器重启后切换显示！",, False)
    End Sub

    '自动更新
    Private Sub ComboSystemActivity_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles ComboSystemActivity.SelectionChanged
        If AniControlEnabled <> 0 Then Return
        If ComboSystemActivity.SelectedIndex <> 2 Then Return
        If MyMsgBox("若选择此项，即使在将来出现严重问题时，你也无法获取相关通知。" & vbCrLf &
                    "例如，如果发现某个版本游戏存在严重 Bug，你可能就会因为无法得到通知而导致无法预知的后果。" & vbCrLf & vbCrLf &
                    "一般选择 仅在有重要通知时显示公告 就可以让你尽量不受打扰了。" & vbCrLf &
                    "除非你在制作服务器整合包，或时常手动更新启动器，否则极度不推荐选择此项！", "警告", "我知道我在做什么", "取消", IsWarn:=True) = 2 Then
            If e.RemovedItems.Count > 0 Then ComboSystemActivity.SelectedItem = e.RemovedItems(0)
        End If
    End Sub
    Private Sub ComboSystemUpdate_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles ComboSystemUpdate.SelectionChanged
        If AniControlEnabled <> 0 Then Return
        If ComboSystemUpdate.SelectedIndex <> 3 Then Return
        If MyMsgBox("若选择此项，即使在启动器将来出现严重问题时，你也无法获取更新并获得修复。" & vbCrLf &
                    "例如，如果官方修改了登录方式，从而导致现有启动器无法登录，你可能就会因为无法更新而无法开始游戏。" & vbCrLf & vbCrLf &
                    "一般选择 仅在有重大漏洞更新时显示提示 就可以让你尽量不受打扰了。" & vbCrLf &
                    "除非你在制作服务器整合包，或时常手动更新启动器，否则极度不推荐选择此项！", "警告", "我知道我在做什么", "取消", IsWarn:=True) = 2 Then
            If e.RemovedItems.Count > 0 Then ComboSystemUpdate.SelectedItem = e.RemovedItems(0)
        End If
    End Sub
    Private Sub BtnSystemUpdate_Click(sender As Object, e As EventArgs) Handles BtnSystemUpdate.Click
        UpdateCheckByButton()
    End Sub
    ''' <summary>
    ''' 启动器是否已经是最新版？
    ''' 若返回 Nothing，则代表无更新缓存文件或出错。
    ''' </summary>
    Public Shared Function IsLauncherNewest() As Boolean?
        Try
            '确认服务器公告是否正常
            Dim ServerContent As String = If(FileUtils.TryReadAsString(PathTemp & "Cache\Notice.cfg"), "")
            If ServerContent.Split("|").Count < 3 Then Return Nothing
            '确认是否为最新
            Return ServerContent.Split("|")(If(BuildType = BuildTypes.Release, 2, 1)) <= VersionCode
        Catch ex As Exception
            Logger.Error(ex, "确认启动器更新失败")
            Return Nothing
        End Try
    End Function

#Region "导出 / 导入设置"

    Private Sub BtnSystemSettingExp_Click(sender As Object, e As MouseButtonEventArgs) Handles BtnSystemSettingExp.Click
        Hint("该功能尚在开发中！")
    End Sub
    Private Sub BtnSystemSettingImp_Click(sender As Object, e As MouseButtonEventArgs) Handles BtnSystemSettingImp.Click
        Hint("该功能尚在开发中！")
    End Sub

#End Region

End Class
