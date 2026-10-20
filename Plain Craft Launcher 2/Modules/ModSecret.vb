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
    '地区判定与下载限制共用缓存：优先使用获得许可的装置位置，随后检查 Windows 地区和网络出口。
    Public Structure RegionStatus
        Public CountryCode As String
        Public DownloadAllowed As Boolean
        Public Reason As String
        Public Details As String
    End Structure

    Private Structure RegionLookup
        Public Source As String
        Public CountryCode As String
        Public Summary As String
    End Structure

    Private ReadOnly RegionDecisionLock As New Object
    Private RegionStatusCache As RegionStatus? = Nothing
    Private RegionCheckExpires As Date = Date.MinValue

    '沿用原来的限制入口；非地区检查版始终不拦截 Minecraft 下载。
    Public ReadOnly Property RestrictMcDownload As Boolean
        Get
            If Not EnableRegionCheck Then Return False
            Return Not GetRegionStatus().DownloadAllowed
        End Get
    End Property

    Private Function IsValidCountryCode(Value As String) As Boolean
        If String.IsNullOrWhiteSpace(Value) OrElse Value.Length <> 2 OrElse Value = "XX" Then Return False
        Return Value.All(Function(Ch) Ch >= "A"c AndAlso Ch <= "Z"c)
    End Function

    'Windows 国家/地区设置是用户可编辑的信号，不单独授予下载权限。
    <System.Runtime.InteropServices.DllImport("kernel32.dll", EntryPoint:="GetUserGeoID")>
    Private Function NativeGetUserGeoID(GeoClass As Integer) As Integer
    End Function

    <System.Runtime.InteropServices.DllImport("kernel32.dll", EntryPoint:="GetGeoInfoW", CharSet:=System.Runtime.InteropServices.CharSet.Unicode)>
    Private Function NativeGetGeoInfo(GeoId As Integer, GeoType As Integer, Value As System.Text.StringBuilder,
                                      ValueLength As Integer, LangId As Integer) As Integer
    End Function

    Private Function GetWindowsRegionCountry() As String
        Try
            Dim GeoId = NativeGetUserGeoID(16) 'GEOCLASS_NATION
            If GeoId <= 0 Then Return Nothing
            Dim Buffer As New System.Text.StringBuilder(16)
            If NativeGetGeoInfo(GeoId, 4, Buffer, Buffer.Capacity, 0) <= 0 Then Return Nothing 'GEO_ISO2
            Dim Country = Buffer.ToString().Trim().ToUpperInvariant()
            Return If(IsValidCountryCode(Country), Country, Nothing)
        Catch
            Return Nothing
        End Try
    End Function

    'Natural Earth 1:110m 公共领域国界数据离线查询，不上传原始经纬度。
    Private Function PointInBorderRing(Lat As Double, Lon As Double, Ring As Newtonsoft.Json.Linq.JArray) As Boolean
        If Ring.Count < 4 Then Return False
        Dim Inside As Boolean = False
        Dim Previous = DirectCast(Ring(Ring.Count - 1), Newtonsoft.Json.Linq.JArray)
        For Each Current As Newtonsoft.Json.Linq.JArray In Ring.Children(Of Newtonsoft.Json.Linq.JArray)()
            Dim X1 = Current(0).Value(Of Double)(), Y1 = Current(1).Value(Of Double)()
            Dim X2 = Previous(0).Value(Of Double)(), Y2 = Previous(1).Value(Of Double)()
            If (Y1 > Lat) <> (Y2 > Lat) Then
                If Lon < (X2 - X1) * (Lat - Y1) / (Y2 - Y1) + X1 Then Inside = Not Inside
            End If
            Previous = Current
        Next
        Return Inside
    End Function

    Private Function CountryAtCoordinate(Lat As Double, Lon As Double, Features As Newtonsoft.Json.Linq.JArray) As String
        For Each Feature As Newtonsoft.Json.Linq.JObject In Features.Children(Of Newtonsoft.Json.Linq.JObject)()
            Dim Code = Feature("c").ToString()
            Dim Polygons = DirectCast(Feature("p"), Newtonsoft.Json.Linq.JArray)
            For Each Polygon As Newtonsoft.Json.Linq.JArray In Polygons.Children(Of Newtonsoft.Json.Linq.JArray)()
                If Polygon.Count = 0 Then Continue For
                If Not PointInBorderRing(Lat, Lon, DirectCast(Polygon(0), Newtonsoft.Json.Linq.JArray)) Then Continue For
                Dim InHole As Boolean = False
                For HoleIndex As Integer = 1 To Polygon.Count - 1
                    If PointInBorderRing(Lat, Lon, DirectCast(Polygon(HoleIndex), Newtonsoft.Json.Linq.JArray)) Then
                        InHole = True
                        Exit For
                    End If
                Next
                If Not InHole Then Return Code
            Next
        Next
        Return Nothing
    End Function

    '只在用户明确同意后请求高精度 Windows 装置定位。
    'Windows 可以使用 GPS、Wi-Fi 等定位技术，不保证实际使用卫星 GPS。
    Private Function GetDeviceCountry() As String
        Try
            Using Watcher As New System.Device.Location.GeoCoordinateWatcher(System.Device.Location.GeoPositionAccuracy.High)
                If Not Watcher.TryStart(False, TimeSpan.FromSeconds(7)) Then Return Nothing
                If Watcher.Permission <> System.Device.Location.GeoPositionPermission.Granted Then Return Nothing
                Dim Coordinate = Watcher.Position.Location
                If Coordinate Is Nothing OrElse Coordinate.IsUnknown Then Return Nothing
                If Double.IsNaN(Coordinate.HorizontalAccuracy) OrElse Coordinate.HorizontalAccuracy > 10000 Then Return Nothing
                Dim Latitude = Coordinate.Latitude, Longitude = Coordinate.Longitude
                If Double.IsNaN(Latitude) OrElse Double.IsNaN(Longitude) OrElse
                   Math.Abs(Latitude) > 90 OrElse Math.Abs(Longitude) > 180 Then Return Nothing

                Dim Assembly = System.Reflection.Assembly.GetExecutingAssembly()
                Using Stream = Assembly.GetManifestResourceStream("PCLRegionBorders")
                    If Stream Is Nothing Then Return Nothing
                    Using Reader As New System.IO.StreamReader(Stream)
                        Dim Root = Newtonsoft.Json.Linq.JObject.Parse(Reader.ReadToEnd())
                        Dim Features = DirectCast(Root("features"), Newtonsoft.Json.Linq.JArray)
                        Dim Country = CountryAtCoordinate(Latitude, Longitude, Features)
                        If Not IsValidCountryCode(Country) Then Return Nothing

                        '邻近边界或离线地图无法确定时回退验证，避免边境定位偏差。
                        Dim Offset = Math.Max(Coordinate.HorizontalAccuracy / 111320.0, 0.04)
                        Dim LonOffset = Offset / Math.Max(Math.Cos(Latitude * Math.PI / 180.0), 0.1)
                        If CountryAtCoordinate(Latitude + Offset, Longitude, Features) <> Country OrElse
                           CountryAtCoordinate(Latitude - Offset, Longitude, Features) <> Country OrElse
                           CountryAtCoordinate(Latitude, Longitude + LonOffset, Features) <> Country OrElse
                           CountryAtCoordinate(Latitude, Longitude - LonOffset, Features) <> Country Then Return Nothing
                        Return Country
                    End Using
                End Using
            End Using
        Catch
            '不记录 GPS 坐标或向网络上传坐标，失败时由后备来源确认。
            Return Nothing
        End Try
    End Function

    '每个查询单独处理失败，避免把定位服务故障误判为 VPN。
    Private Function LookupRegion(Source As String, Url As String, Optional Field As String = Nothing) As RegionLookup
        Dim Result As New RegionLookup With {.Source = Source, .CountryCode = Nothing, .Summary = Source & "：查询失败或超时"}
        Try
            Using Client As New HttpClient()
                Client.Timeout = TimeSpan.FromSeconds(6)
                Client.DefaultRequestHeaders.UserAgent.ParseAdd("PCL2-Ray-Region/2.0")
                Dim Text = Client.GetStringAsync(Url).GetAwaiter().GetResult()
                Dim Country As String = Nothing
                If Source = "Cloudflare" Then
                    For Each Line In Text.Split({vbLf}, StringSplitOptions.RemoveEmptyEntries)
                        If Line.StartsWith("loc=", StringComparison.OrdinalIgnoreCase) Then
                            Country = Line.Substring(4).Trim().ToUpperInvariant()
                            Exit For
                        End If
                    Next
                Else
                    Dim Json = Newtonsoft.Json.Linq.JObject.Parse(Text)
                    If Source = "ipwho.is" AndAlso
                       (Json("success") Is Nothing OrElse Not Json("success").ToObject(Of Boolean)) Then
                        Result.Summary = Source & "：服务未返回有效查询结果"
                        Return Result
                    End If
                    Country = If(Json(Field)?.ToString(), "").Trim().ToUpperInvariant()
                End If
                If IsValidCountryCode(Country) Then
                    Result.CountryCode = Country
                    Result.Summary = Source & "：" & Country
                Else
                    Result.Summary = Source & "：未返回有效国家代码"
                End If
            End Using
        Catch ex As Exception
            '只输出服务名与诊断分类，不打印原始响应或公网 IP。
            Result.Summary = Source & "：连接失败或超时"
        End Try
        Return Result
    End Function

    '代理/VPN 线索仅供解释网络环境；Windows 默认代理和虚拟网卡不证明用户使用 VPN。
    Private Function GetNetworkHint() As String
        Try
            Dim Hints As New List(Of String)
            Dim Target As New Uri("https://www.cloudflare.com/")
            Dim Proxy = System.Net.WebRequest.DefaultWebProxy
            If Proxy IsNot Nothing Then
                Dim Route = Proxy.GetProxy(Target)
                If Route IsNot Nothing AndAlso Not Route.Equals(Target) Then
                    Hints.Add("系统配置了代理路由")
                End If
            End If
            For Each Adapter In System.Net.NetworkInformation.NetworkInterface.GetAllNetworkInterfaces()
                If Adapter.OperationalStatus <> System.Net.NetworkInformation.OperationalStatus.Up Then Continue For
                Dim Name = (Adapter.Name & " " & Adapter.Description).ToLowerInvariant()
                If {"wireguard", "openvpn", "wintun", "tailscale", "zerotier", "softether"}.Any(Function(X) Name.Contains(X)) Then
                    Hints.Add("存在已启用的隧道网卡")
                    Exit For
                End If
            Next
            Return If(Hints.Count = 0, "未发现明确的代理线索（不代表不存在 VPN）",
                String.Join("；", Hints) & "（仅供参考，未直接作为封锁依据）")
        Catch ex As Exception
            Return "代理检查不可用（不会被当作 VPN 证据）"
        End Try
    End Function

    '至少两个定位源一致才视为已核实；第三个来源仅在前两个失败或冲突时启用。
    '任一有效来源明确返回 CN 时保守拒绝，以免出现互相冲突却放行的情况。
    Public Function GetRegionStatus(Optional ForceRefresh As Boolean = False) As RegionStatus
        SyncLock RegionDecisionLock
            If Not ForceRefresh AndAlso RegionStatusCache.HasValue AndAlso Date.UtcNow < RegionCheckExpires Then
                Return RegionStatusCache.Value
            End If

            '设备定位优先。仅在使用者主动授权后访问 Windows 设备位置。
            Dim ConsentValue = RegistryUtils.TryRead("HKEY_CURRENT_USER\Software\PCLRegionUi", "AllowDeviceLocationV1", Nothing)
            Dim DeviceCountry As String = Nothing
            If ConsentValue IsNot Nothing AndAlso ConsentValue.ToString() = "1" Then
                DeviceCountry = GetDeviceCountry()
            End If
            If IsValidCountryCode(DeviceCountry) Then
                Dim DeviceAllowed = Not EnableRegionCheck OrElse DeviceCountry <> "CN"
                Dim DeviceReason = If(DeviceAllowed, "当前地区可以使用 Minecraft 下载。", "当前地区暂不支持 Minecraft 下载。")
                Dim DeviceStatus As New RegionStatus With {
                    .CountryCode = DeviceCountry, .DownloadAllowed = DeviceAllowed,
                    .Reason = DeviceReason, .Details = "地区信息已更新。"
                }
                RegionStatusCache = DeviceStatus
                RegionCheckExpires = Date.UtcNow.AddMinutes(5)
                Return DeviceStatus
            End If

            'Windows 地区先取得并供判断参考；由于它可手动修改，下载放行仍需 IP 确认。
            Dim WindowsCountry = GetWindowsRegionCountry()
            Dim First = System.Threading.Tasks.Task.Run(Function() LookupRegion(
                "Cloudflare", "https://www.cloudflare.com/cdn-cgi/trace"))
            Dim Second = System.Threading.Tasks.Task.Run(Function() LookupRegion(
                "ipwho.is", "https://ipwho.is/", "country_code"))
            Dim Results As New List(Of RegionLookup) From {First.Result, Second.Result}

            If Not IsValidCountryCode(Results(0).CountryCode) OrElse
               Not IsValidCountryCode(Results(1).CountryCode) OrElse
               Results(0).CountryCode <> Results(1).CountryCode Then
                Results.Add(LookupRegion("country.is", "https://api.country.is/", "country"))
            End If

            Dim ValidCountries = Results.Where(Function(R) IsValidCountryCode(R.CountryCode)).
                Select(Function(R) R.CountryCode).ToList()
            Dim AnyMainlandChina As Boolean = ValidCountries.Contains("CN")
            Dim ConfirmedCode As String = Nothing
            For Each Candidate In ValidCountries.Distinct()
                If ValidCountries.Where(Function(C) C = Candidate).Count() >= 2 Then
                    ConfirmedCode = Candidate
                    Exit For
                End If
            Next

            Dim DisplayFallback As String = If(ValidCountries.Count = 1, ValidCountries(0) & "（待核实）",
                If(IsValidCountryCode(WindowsCountry), WindowsCountry & "（待核实）", "未知"))
            Dim CountryCode = If(AnyMainlandChina, "CN", If(ConfirmedCode IsNot Nothing, ConfirmedCode, DisplayFallback))
            '由于 Windows 国家/地区可手动修改，不用它单独解锁下载。
            Dim Allowed As Boolean = Not EnableRegionCheck OrElse
                (Not AnyMainlandChina AndAlso ConfirmedCode IsNot Nothing AndAlso
                 ValidCountries.All(Function(C) C = ConfirmedCode))
            Dim Reason As String
            If Not EnableRegionCheck Then
                Reason = "当前版本可以使用 Minecraft 下载。"
            ElseIf AnyMainlandChina Then
                Reason = "当前地区暂不支持 Minecraft 下载。"
            ElseIf Allowed Then
                Reason = "当前地区可以使用 Minecraft 下载。"
            Else
                Reason = "暂时无法确认下载服务的适用地区，请稍后重试。"
            End If
            Dim Status As New RegionStatus With {
                .CountryCode = CountryCode, .DownloadAllowed = Allowed,
                .Reason = Reason, .Details = "地区信息已更新。"
            }
            RegionStatusCache = Status
            RegionCheckExpires = Date.UtcNow.AddMinutes(5)
            Return Status
        End SyncLock
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
