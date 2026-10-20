'由于包含加解密等安全信息，本文件中的部分代码已被删除

Friend Module ModSecret

    '标注 PCL 的不同分支，仅用于替换标记
    Public Const VersionBranchMain As String = "OpenSource"
    '在开源版的注册表与常规版的注册表隔离，以防数据冲突
    '用于保存选择的独立注册表路径，避免随隔离模式切换而丢失选择。
    Private Const RegistryChoicePath As String = "HKEY_CURRENT_USER\Software\PCLRegistryChoice"
    Private _registryIsolated As Boolean = True

    Public ReadOnly Property RegFolder As String
        Get
            Return If(_registryIsolated, "PCLDebug", "PCL")
        End Get
    End Property

    Public ReadOnly Property RegistryIsolated As Boolean
        Get
            Return _registryIsolated
        End Get
    End Property

    '只接受明确存储的 0 / 1，未选择或无效时继续弹出首次选择。
    Friend Function GetRegistryIsolationChoice() As Boolean?
        Dim Saved = RegistryUtils.TryRead(RegistryChoicePath, "Isolated", Nothing)
        If Saved Is Nothing Then Return Nothing
        Select Case Saved.ToString()
            Case "1"
                Return True
            Case "0"
                Return False
            Case Else
                Return Nothing
        End Select
    End Function

    Friend Function TrySaveRegistryIsolationChoice(Isolate As Boolean) As Boolean
        Try
            RegistryUtils.Write(RegistryChoicePath, "Isolated", If(Isolate, "1", "0"))
            Dim Saved = GetRegistryIsolationChoice()
            If Saved.HasValue AndAlso Saved.Value = Isolate Then Return True
            Throw New InvalidOperationException("无法确认注册表选项已正确保存。")
        Catch ex As Exception
            System.Windows.MessageBox.Show("保存注册表隔离选项失败：" & ex.Message & vbCrLf &
                "本次不会切换设置，请检查当前用户的注册表写入权限。", "保存失败",
                System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error)
            Return False
        End Try
    End Function

    '必须先决定注册表目录，才能加载启动器设置。
    Friend Function InitializeRegistryIsolation() As Boolean
        Dim Choice = GetRegistryIsolationChoice()
        If Not Choice.HasValue Then
            Dim Answer = System.Windows.MessageBox.Show(
                "是否要启用注册表隔离？" & vbCrLf & vbCrLf &
                "是：使用开源版独立设置（PCLDebug）。" & vbCrLf &
                "否：与 PCL 正式版共用设置（PCL）；对设置的更改也会影响正式版。" & vbCrLf & vbCrLf &
                "选择只需进行一次，之后可以在 设置 → 其他 中修改。" & vbCrLf &
                "取消或关闭此窗口将退出启动器，下次启动会再次询问。",
                "注册表隔离", System.Windows.MessageBoxButton.YesNoCancel,
                System.Windows.MessageBoxImage.Question)
            If Answer <> System.Windows.MessageBoxResult.Yes AndAlso
               Answer <> System.Windows.MessageBoxResult.No Then Return False
            Dim Isolate = Answer = System.Windows.MessageBoxResult.Yes
            If Not TrySaveRegistryIsolationChoice(Isolate) Then Return False
            Choice = Isolate
        End If
        _registryIsolated = Choice.Value
        Return True
    End Function
    '用于微软登录的 ClientId
    Public OAuthClientId As String = If(Environment.GetEnvironmentVariable("PCL_MS_CLIENT_ID"), "")
    'CurseForge API Key
    Public CurseForgeAPIKey As String = If(Environment.GetEnvironmentVariable("PCL_CURSEFORGE_API_KEY"), "")
    '用于匿名数据收集的腾讯云日志服务上报 URL，形如 https://{region}.cls.tencentcs.com/track?topic_id={topic_id}
    Public Const ClsBaseUrl As String = ""

#Region "网络鉴权"

    '保留 CurseForge 旧实现，默认只使用 Modrinth。
    Public Const EnableCurseForge As Boolean = False
    '公开版启用地理区域限制；私有构建可以独立覆盖此常量。
    Public Const EnableRegionCheck As Boolean = True
    '地区判定与下载限制共用相同的缓存，避免显示的状态与实际下载政策不一致。
    Public Structure RegionStatus
        Public CountryCode As String
        Public DownloadAllowed As Boolean
        Public Reason As String
    End Structure
    Private ReadOnly RegionDecisionLock As New Object
    Private RegionStatusCache As RegionStatus? = Nothing
    Private RegionCheckExpires As Date = Date.MinValue

    '沿用 RestrictMcDownload 的所有既有限制入口。
    Public ReadOnly Property RestrictMcDownload As Boolean
        Get
            If Not EnableRegionCheck Then Return False
            Return Not GetRegionStatus().DownloadAllowed
        End Get
    End Property

    '仅查询网络出口国家代码，不申请 GPS 定位或保存精确位置。
    Public Function GetRegionStatus(Optional ForceRefresh As Boolean = False) As RegionStatus
        SyncLock RegionDecisionLock
            If Not ForceRefresh AndAlso RegionStatusCache.HasValue AndAlso Date.UtcNow < RegionCheckExpires Then
                Return RegionStatusCache.Value
            End If

            Dim Country As String = "未知"
            Dim IsConfirmed As Boolean = False
            Dim CheckError As String = Nothing
            Try
                Using Client As New HttpClient()
                    Client.Timeout = TimeSpan.FromSeconds(8)
                    Client.DefaultRequestHeaders.UserAgent.ParseAdd("Ray20123315/PCL-RegionCheck/1.1")
                    Dim Trace = Client.GetStringAsync("https://www.cloudflare.com/cdn-cgi/trace").GetAwaiter().GetResult()
                    For Each Line In Trace.Split({vbLf}, StringSplitOptions.RemoveEmptyEntries)
                        If Line.StartsWith("loc=", StringComparison.OrdinalIgnoreCase) Then
                            Dim Value = Line.Substring(4).Trim().ToUpperInvariant()
                            If Value.Length = 2 AndAlso Value <> "XX" Then Country = Value
                            Exit For
                        End If
                    Next
                    Dim Result = Newtonsoft.Json.Linq.JObject.Parse(
                        Client.GetStringAsync("https://ipwho.is/").GetAwaiter().GetResult())
                    Dim Country2 As String = If(Result("country_code")?.ToString(), "").ToUpperInvariant()
                    IsConfirmed = Country <> "未知" AndAlso Country2.Length = 2 AndAlso Country = Country2 AndAlso
                        Result("success") IsNot Nothing AndAlso Result("success").ToObject(Of Boolean)
                    If Not IsConfirmed Then CheckError = "两个 IP 定位来源未能一致确认国家或地区。"
                End Using
            Catch ex As Exception
                CheckError = "无法取得完整的 IP 地区信息。"
            End Try

            Dim ProxyDetected As Boolean = If(EnableRegionCheck, HasKnownProxyOrVpn(), False)
            Dim Allowed As Boolean = Not EnableRegionCheck OrElse
                (IsConfirmed AndAlso Country <> "CN" AndAlso Not ProxyDetected)
            Dim Reason As String
            If Not EnableRegionCheck Then
                Reason = "此版本未启用地区下载限制，IP 地区信息仅供参考。"
            ElseIf Country = "CN" Then
                Reason = "检测到中国大陆网络出口，Minecraft 下载已禁用。"
            ElseIf ProxyDetected Then
                Reason = "检测到可能影响地区判定的代理或 VPN，下载已暂停。"
            ElseIf Not IsConfirmed Then
                Reason = If(CheckError, "地区尚未得到确认，下载已暂停。")
            Else
                Reason = "地区检查通过，可以使用 Minecraft 下载功能。"
            End If
            Dim Status As New RegionStatus With {
                .CountryCode = Country, .DownloadAllowed = Allowed, .Reason = Reason
            }
            RegionStatusCache = Status
            RegionCheckExpires = Date.UtcNow.AddMinutes(2)
            Return Status
        End SyncLock
    End Function

    '只能识别一部分代理及 VPN，无法保证不存在未识别的网络隧道。
    Private Function HasKnownProxyOrVpn() As Boolean
        Try
            Dim Proxy = System.Net.WebRequest.DefaultWebProxy
            If Proxy IsNot Nothing AndAlso Not Proxy.IsBypassed(New Uri("https://www.cloudflare.com")) Then Return True
            For Each Adapter In System.Net.NetworkInformation.NetworkInterface.GetAllNetworkInterfaces()
                If Adapter.OperationalStatus <> System.Net.NetworkInformation.OperationalStatus.Up Then Continue For
                If Adapter.NetworkInterfaceType = System.Net.NetworkInformation.NetworkInterfaceType.Tunnel Then Return True
                Dim Info = (Adapter.Name & " " & Adapter.Description).ToLowerInvariant()
                If {"wireguard", "openvpn", "wintun", "tailscale", "zerotier", "softether", "tap-windows", "vpn", "tun "}.
                    Any(Function(Marker) Info.Contains(Marker)) Then Return True
            Next
        Catch
            Return True
        End Try
        Return False
    End Function
    Friend Function SecretCdnSign(UrlWithMark As String) As String
        If Not UrlWithMark.EndsWithF("{CDN}") Then Return UrlWithMark
        Return UrlWithMark.Replace("{CDN}", "").Replace(" ", "%20")
    End Function
    ''' <summary>
    ''' 设置 Headers 的 UA、Referer。
    ''' </summary>
    Friend Sub SecretHeadersSign(Url As String, ByRef Req As HttpRequestMessage, Optional SimulateBrowserHeaders As Boolean = False)
        If Not Req.Headers.UserAgent.Any Then
            If Url.Contains("baidupcs.com") OrElse Url.Contains("baidu.com") Then
                Req.Headers.Add("User-Agent", "LogStatistic")  '#4951
            ElseIf SimulateBrowserHeaders Then
                Req.Headers.Add("User-Agent", $"PCL2/{VersionBaseName}.{CInt(BuildType)} Mozilla/5.0 AppleWebKit/537.36 Chrome/63.0.3239.132 Safari/537.36")
            Else
                Req.Headers.Add("User-Agent", $"PCL2/{VersionBaseName}.{CInt(BuildType)}")
            End If
        End If
        If Not SimulateBrowserHeaders Then Req.Headers.Add("Referer", $"http://{VersionCode}.open.pcl2.server/")
        If EnableCurseForge AndAlso (Url.Contains("api.curseforge.com") OrElse Url.Contains("forgecdn.net")) Then Req.Headers.Add("x-api-key", CurseForgeAPIKey)
    End Sub

#End Region

#Region "主题"

    Public Color1 As New MyColor(52, 61, 74)
    Public Color2 As New MyColor(11, 91, 203)
    Public Color3 As New MyColor(19, 112, 243)
    Public Color4 As New MyColor(72, 144, 245)
    Public Color5 As New MyColor(150, 192, 249)
    Public Color6 As New MyColor(213, 230, 253)
    Public Color7 As New MyColor(222, 236, 253)
    Public Color8 As New MyColor(234, 242, 254)
    Public ColorBg0 As New MyColor(150, 192, 249)
    Public ColorBg1 As New MyColor(190, Color7)
    Public ColorGray1 As New MyColor(64, 64, 64)
    Public ColorGray2 As New MyColor(115, 115, 115)
    Public ColorGray3 As New MyColor(140, 140, 140)
    Public ColorGray4 As New MyColor(166, 166, 166)
    Public ColorGray5 As New MyColor(204, 204, 204)
    Public ColorGray6 As New MyColor(235, 235, 235)
    Public ColorGray7 As New MyColor(240, 240, 240)
    Public ColorGray8 As New MyColor(245, 245, 245)
    Public ColorSemiTransparent As New MyColor(1, Color8)

    Public ThemeNow As Integer = -1
    Public ColorHue As Integer = 210, ColorSat As Integer = 85, ColorLightAdjust As Integer = 0, ColorHueTopbarDelta As OneOf(Of Integer, Integer()) = 0
    Public ThemeDontClick As Integer = 0

    Public Sub ThemeRefresh(Optional NewTheme As Integer = -1)
        Try
            If ThemeNow = NewTheme AndAlso NewTheme >= 0 Then Return
            If NewTheme >= 0 Then ThemeNow = NewTheme

            Color1 = New MyColor().FromHSL2(ColorHue, ColorSat * 0.2, 25 + ColorLightAdjust * 0.3)
            Color2 = New MyColor().FromHSL2(ColorHue, ColorSat, 45 + ColorLightAdjust)
            Color3 = New MyColor().FromHSL2(ColorHue, ColorSat, 55 + ColorLightAdjust)
            Color4 = New MyColor().FromHSL2(ColorHue, ColorSat, 65 + ColorLightAdjust)
            Color5 = New MyColor().FromHSL2(ColorHue, ColorSat, 80 + ColorLightAdjust * 0.4)
            Color6 = New MyColor().FromHSL2(ColorHue, ColorSat, 91 + ColorLightAdjust * 0.1)
            Color7 = New MyColor().FromHSL2(ColorHue, ColorSat, 95)
            Color8 = New MyColor().FromHSL2(ColorHue, ColorSat, 97)
            ColorBg0 = Color4 * 0.4 + Color5 * 0.4 + ColorGray4 * 0.2
            ColorBg1 = New MyColor(190, Color7)

            ColorSemiTransparent = New MyColor(1, Color8)
            Application.Current.Resources("ColorBrush1") = New SolidColorBrush(Color1)
            Application.Current.Resources("ColorBrush2") = New SolidColorBrush(Color2)
            Application.Current.Resources("ColorBrush3") = New SolidColorBrush(Color3)
            Application.Current.Resources("ColorBrush4") = New SolidColorBrush(Color4)
            Application.Current.Resources("ColorBrush5") = New SolidColorBrush(Color5)
            Application.Current.Resources("ColorBrush6") = New SolidColorBrush(Color6)
            Application.Current.Resources("ColorBrush7") = New SolidColorBrush(Color7)
            Application.Current.Resources("ColorBrush8") = New SolidColorBrush(Color8)
            Application.Current.Resources("ColorBrushBg0") = New SolidColorBrush(ColorBg0)
            Application.Current.Resources("ColorBrushBg1") = New SolidColorBrush(ColorBg1)
            Application.Current.Resources("ColorObject1") = CType(Color1, Color)
            Application.Current.Resources("ColorObject2") = CType(Color2, Color)
            Application.Current.Resources("ColorObject3") = CType(Color3, Color)
            Application.Current.Resources("ColorObject4") = CType(Color4, Color)
            Application.Current.Resources("ColorObject5") = CType(Color5, Color)
            Application.Current.Resources("ColorObject6") = CType(Color6, Color)
            Application.Current.Resources("ColorObject7") = CType(Color7, Color)
            Application.Current.Resources("ColorObject8") = CType(Color8, Color)
            Application.Current.Resources("ColorObjectBg0") = CType(ColorBg0, Color)
            Application.Current.Resources("ColorObjectBg1") = CType(ColorBg1, Color)
            ThemeRefreshMain()
        Catch ex As Exception
            Logger.Error(ex, "刷新主题颜色失败", LogBehavior.Toast)
        End Try
    End Sub
    Public Sub ThemeRefreshMain()
        RunInUi(
        Sub()
            If Not FrmMain.IsLoaded Then Return
            '顶部条背景
            Dim Brush = New LinearGradientBrush With {.EndPoint = New Point(1, 0), .StartPoint = New Point(0, 0)}
            Dim Deltas = ColorHueTopbarDelta.Switch(Function(d) New Integer() {-d, 0, d}, Function(d) d)
            Brush.GradientStops.Add(New GradientStop With {.Offset = 0, .Color = New MyColor().FromHSL2(ColorHue + Deltas(0), ColorSat, 48 + ColorLightAdjust)})
            Brush.GradientStops.Add(New GradientStop With {.Offset = 0.5, .Color = New MyColor().FromHSL2(ColorHue + Deltas(1), ColorSat, 54 + ColorLightAdjust)})
            Brush.GradientStops.Add(New GradientStop With {.Offset = 1, .Color = New MyColor().FromHSL2(ColorHue + Deltas(2), ColorSat, 48 + ColorLightAdjust)})
            FrmMain.PanTitle.Background = Brush
            FrmMain.PanTitle.Background.Freeze()
            '主页面背景
            If Settings.Get(Of Boolean)("UiBackgroundColorful") Then
                Brush = New LinearGradientBrush With {.EndPoint = New Point(0.1, 1), .StartPoint = New Point(0.9, 0)}
                Brush.GradientStops.Add(New GradientStop With {.Offset = -0.1, .Color = New MyColor().FromHSL2(ColorHue - 15, ColorSat * 0.8, 91)})
                Brush.GradientStops.Add(New GradientStop With {.Offset = 0.4, .Color = New MyColor().FromHSL2(ColorHue, ColorSat * 0.8, 91)})
                Brush.GradientStops.Add(New GradientStop With {.Offset = 1.1, .Color = New MyColor().FromHSL2(ColorHue + 15, ColorSat * 0.8, 91)})
                FrmMain.PanForm.Background = Brush
            Else
                FrmMain.PanForm.Background = New MyColor(245, 245, 245)
            End If
            FrmMain.PanForm.Background.Freeze()
        End Sub)
    End Sub
    Friend Sub ThemeCheckAll(EffectSetup As Boolean)
    End Sub
    Friend Function ThemeCheckOne(Id As Integer) As Boolean
        Return True
    End Function
    Friend Function ThemeUnlock(Id As Integer, Optional ShowDoubleHint As Boolean = True, Optional UnlockHint As String = Nothing) As Boolean
        Return False
    End Function

#End Region

#Region "更新"

    Friend Sub UpdateCheckByButton()
        Hint("该版本中不包含更新功能……")
    End Sub

    Friend IsUpdateWaitingRestart As Boolean = False
    Public Sub UpdateRestart(TriggerRestartAndByEnd As Boolean)
    End Sub
    Public Sub UpdateReplace(ProcessId As Integer, OldFileName As String, NewFileName As String, TriggerRestart As Boolean)
    End Sub

    ''' <summary>
    ''' 确保 PathTemp/Latest.exe 是最新正式版的 PCL，它会被用于整合包打包。
    ''' 如果不是，则下载一个。
    ''' </summary>
    Friend Sub DownloadLatestPCL(Optional LoaderToSyncProgress As LoaderBase = Nothing)
        '注意：如果要自行实现这个功能，请换用另一个文件路径，以免与官方版本冲突
    End Sub

#End Region

#Region "联网配置"

    ''' <summary>
    ''' 联网获取的配置信息。
    ''' 若获取失败或仍在获取中，可能为 Nothing。
    ''' </summary>
    Public ServerConfig As JObject

    Public ServerLoader As New LoaderTask(Of Integer, Integer)("PCL 配置更新", Sub() Logger.Info("该版本中不包含更新通知功能……"), Priority:=ThreadPriority.BelowNormal) With
        {.ReloadTimeout = 1000 * 60 * 60} '超时 1 小时

#End Region

#Region "赞助等级"

    Public ReadOnly Property CurrentRank As DonationRank
        Get
            Return DonationRank.None
        End Get
    End Property

    Public Sub InputPotatoCode(IsUpdating As Boolean)
    End Sub
    Friend Sub GeneratePotatoCode()
    End Sub

    ''' <summary>
    ''' 获取设备识别码。
    ''' </summary>
    Friend Function GetIdentify() As String
        Return "0000-0000-0000-0000"
    End Function

#End Region

End Module
