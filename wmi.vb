'Private nic_properties() As String = New String() {"ArpAlwaysSourceRoute", "ArpUseEtherSNAP", "Caption", "DatabasePath", "DeadGWDetectEnabled", "DefaultIPGateway", "DefaultTOS", "DefaultTTL", "Description", "DHCPEnabled", "DHCPLeaseExpires", "DHCPLeaseObtained", "DHCPServer", "DNSDomain", "DNSDomainSuffixSearchOrder", "DNSEnabledForWINSResolution", "DNSHostName", "DNSServerSearchOrder", "DomainDNSRegistrationEnabled", "ForwardBufferMemory", "FullDNSRegistrationEnabled", "GatewayCostMetric", "IGMPLevel", "Index", "InterfaceIndex", "IPAddress", "IPConnectionMetric", "IPEnabled", "IPFilterSecurityEnabled", "IPPortSecurityEnabled", "IPSecPermitIPProtocols", "IPSecPermitTCPPorts", "IPSecPermitUDPPorts", "IPSubnet", "IPUseZeroBroadcast", "IPXAddress", "IPXEnabled", "IPXFrameType", "IPXMediaType", "IPXNetworkNumber", "IPXVirtualNetNumber", "KeepAliveInterval", "KeepAliveTime", "MACAddress", "MTU", "NumForwardPackets", "PMTUBHDetectEnabled", "PMTUDiscoveryEnabled", "ServiceName", "SettingID", "TcpipNetbiosOptions", "TcpMaxConnectRetransmissions", "TcpMaxDataRetransmissions", "TcpNumConnections", "TcpUseRFC1122UrgentPointer", "TcpWindowSize", "WINSEnableLMHostsLookup", "WINSHostLookupFile", "WINSPrimaryServer", "WINSScopeID", "WINSSecondaryServer"}

Imports System.Management
Friend Class wmiObject
    Implements IDisposable
    Private MgmtClass As ManagementClass
    Private _nics As List(Of wmiNICObject)
    Private disposedValue As Boolean
    Sub New()
        Create()
    End Sub
    Private Sub Create()
        Try
            Me.MgmtClass = New ManagementClass("Win32_NetworkAdapterConfiguration")
            Me._nics = New List(Of wmiNICObject)
            For Each objMO As ManagementObject In MgmtClass.GetInstances
                If objMO("IPEnabled").ToString.Equals("True") Then
                    'Debug.WriteLine(objMO("NetConnectionStatus"))
                    _nics.Add(New wmiNICObject(objMO))
                End If
            Next
        Catch
        End Try
    End Sub

    Public Sub Reset()
        If _nics IsNot Nothing Then
            For Each n In _nics
                n.Dispose()
            Next
            _nics.Clear()
        End If
        If MgmtClass IsNot Nothing Then MgmtClass.Dispose()
        Me.Create()
    End Sub
    Public ReadOnly Property IsValid As Boolean
        Get
            If disposedValue Then Return False
            Return MgmtClass IsNot Nothing
        End Get
    End Property
    Public ReadOnly Property NICs As List(Of wmiNICObject)
        Get
            Return _nics
        End Get
    End Property

    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                If _nics IsNot Nothing Then
                    For Each n In _nics
                        n.Dispose()
                    Next
                    _nics.Clear()
                End If
                If MgmtClass IsNot Nothing Then MgmtClass.Dispose()
            End If
            _nics = Nothing
            MgmtClass = Nothing
        End If
        Me.disposedValue = True
    End Sub
    Public Sub Dispose() Implements IDisposable.Dispose
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub


End Class

Friend Class wmiNICObject
    Implements IDisposable
    Private wmiO As ManagementObject
    Private disposedValue As Boolean
    Sub New(MgmtObj As ManagementObject)
        wmiO = MgmtObj
    End Sub
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                If wmiO IsNot Nothing Then wmiO.Dispose()
            End If
            wmiO = Nothing
        End If
        Me.disposedValue = True
    End Sub
    Public Sub Dispose() Implements IDisposable.Dispose
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub

    Public Property IP As String
        Get
            If wmiO Is Nothing OrElse wmiO("IPAddress") Is Nothing OrElse wmiO("IPSubnet") Is Nothing Then Return String.Empty
            Dim ips() As String = CType(wmiO("IPAddress"), String())
            Dim subs() As String = CType(wmiO("IPSubnet"), String())
            Dim retval As New System.Text.StringBuilder
            Dim addr As IPAddress
            For it As Integer = 0 To ips.GetUpperBound(0)
                If it < subs.Length Then
                    If IPAddress.TryParse(ips(it), subs(it), addr) Then
                        retval.Append(", " & addr.ToString)
                    End If
                Else
                    Exit For
                End If
            Next
            Erase ips : Erase subs
            If retval.Length > 0 Then
                If GetString(wmiO("DHCPEnabled")).Equals("True") Then retval.Append(" [DHCP]")
                Return retval.ToString.Substring(2)
            End If
            Return String.Empty
        End Get

        Set(value As String)
            If String.IsNullOrEmpty(value) Then Exit Property
            If value.Equals("dhcp", StringComparison.InvariantCultureIgnoreCase) Or value.Equals("dynamic", StringComparison.InvariantCultureIgnoreCase) Then
                Try
                    wmiO.InvokeMethod("EnableDHCP", Nothing)
                Catch
                End Try
            Else
                Dim addrs As New List(Of IPAddress), addr As IPAddress
                For Each str_addr As String In value.Split(","c)
                    If IPAddress.TryParse(str_addr.Trim, addr) Then addrs.Add(addr)
                Next
                If addrs.Count = 0 Then Exit Property

                Dim ips As New List(Of String), subnets As New List(Of String)
                For Each addr In addrs
                    ips.Add(addr.Address)
                    subnets.Add(addr.SubnetMask)
                Next
                Dim newIP As ManagementBaseObject = Nothing
                Try
                    newIP = wmiO.GetMethodParameters("EnableStatic")
                    newIP("IPAddress") = ips.ToArray
                    newIP("SubnetMask") = subnets.ToArray
                    wmiO.InvokeMethod("EnableStatic", newIP, Nothing)
                Catch
                Finally
                    If newIP IsNot Nothing Then newIP.Dispose() : newIP = Nothing
                End Try
                ips.Clear() : ips = Nothing
                subnets.Clear() : subnets = Nothing
                addrs.Clear() : addrs = Nothing
                addr = Nothing
            End If
        End Set
    End Property

    Public Property Gateway As String
        Get
            Return GetString(wmiO("DefaultIPGateway"))
        End Get
        Set(value As String)
            If String.IsNullOrEmpty(value) Then Exit Property
            Dim addr As IPAddress = Nothing
            If IPAddress.TryParse(value.Trim, addr) = False Then Exit Property

            Dim newGateway As ManagementBaseObject = Nothing
            Try
                newGateway = wmiO.GetMethodParameters("SetGateways")
                newGateway("DefaultIPGateway") = New String() {addr.Address}
                newGateway("GatewayCostMetric") = New Integer() {1}
                wmiO.InvokeMethod("SetGateways", newGateway, Nothing)
            Catch
            Finally
                If newGateway IsNot Nothing Then newGateway.Dispose()
                newGateway = Nothing
            End Try
            addr = Nothing
        End Set
    End Property

    Public Property DNS As String
        Get
            Return GetString(wmiO("DNSServerSearchOrder"))
        End Get
        Set(value As String)
            If String.IsNullOrEmpty(value) Then Exit Property
            Dim addr As IPAddress = Nothing, ips As New List(Of String)
            For Each str_addr In value.Split(","c)
                If IPAddress.TryParse(str_addr.Trim, addr) Then ips.Add(addr.Address)
            Next
            If ips.Count = 0 Then Exit Property
            Dim newDNS As ManagementBaseObject = Nothing
            Try
                newDNS = wmiO.GetMethodParameters("SetDNSServerSearchOrder")
                newDNS("DNSServerSearchOrder") = ips.ToArray
                wmiO.InvokeMethod("SetDNSServerSearchOrder", newDNS, Nothing)
            Catch
            Finally
                If newDNS IsNot Nothing Then newDNS.Dispose()
                newDNS = Nothing
            End Try
            ips.Clear() : ips = Nothing
            addr = Nothing
        End Set
    End Property

    Public Property WINS As String
        Get
            If wmiO Is Nothing OrElse wmiO("WINSPrimaryServer") Is Nothing Then Return String.Empty
            Return GetString(wmiO("WINSPrimaryServer")) & _
                If(wmiO("WINSSecondaryServer") IsNot Nothing, ", " & GetString(wmiO("WINSSecondaryServer")), String.Empty)
        End Get
        Set(value As String)
            If String.IsNullOrEmpty(value) Then Exit Property
            Dim addr As IPAddress = Nothing, winses As New List(Of String)
            For Each str_addr In value.Split(","c)
                If IPAddress.TryParse(str_addr.Trim, addr) Then winses.Add(addr.Address)
            Next
            If winses.Count = 0 Then Exit Property

            Dim setWins As ManagementBaseObject = Nothing
            Try
                setWins = wmiO.GetMethodParameters("SetWINSServer")
                setWins.SetPropertyValue("WINSPrimaryServer", winses(0))
                If winses.Count > 1 Then setWins.SetPropertyValue("WINSSecondaryServer", winses(1))

                wmiO.InvokeMethod("SetWINSServer", setWins, Nothing)
            Catch
                If setWins IsNot Nothing Then setWins.Dispose() : setWins = Nothing
            End Try
            winses.Clear() : winses = Nothing
            addr = Nothing
        End Set
    End Property
    Public ReadOnly Property Name As String
        Get
            Return GetString(wmiO("Description"))
        End Get
    End Property

    Private Shared Function GetString(O As Object) As String
        If O Is Nothing Then Return String.Empty
        If IsArray(O) Then
            Dim tb As New System.Text.StringBuilder
            For Each s In CType(O, Array)
                tb.Append(", " & GetString(s))
            Next
            If tb.Length < 2 Then Return tb.ToString
            Return tb.ToString.Substring(2)
        End If
        Return O.ToString
    End Function

    Private Structure IPAddress
        Private _oct0 As Byte, _oct1 As Byte, _oct2 As Byte, _oct3 As Byte, _mask As Byte
        Public ReadOnly Property Address As String
            Get
                Return String.Format("{0}.{1}.{2}.{3}", _oct0, _oct1, _oct2, _oct3)
            End Get
        End Property
        Public ReadOnly Property SubnetMask As String
            Get
                If _mask = 0 Or _mask > 32 Then Return String.Empty
                Dim int_mask As Integer = &HFFFFFFFF Xor CInt(2 ^ (32 - _mask) - 1)
                Return String.Format("{0}.{1}.{2}.{3}", _
                                     ((int_mask >> 24) And &HFF).ToString, _
                                     ((int_mask >> 16) And &HFF).ToString, _
                                     ((int_mask >> 8) And &HFF).ToString, _
                                     (int_mask And &HFF).ToString)
            End Get
        End Property
        Public Overrides Function ToString() As String
            Return String.Format("{0}.{1}.{2}.{3}{4}", _oct0, _oct1, _oct2, _oct3, _
                                 If(_mask > 0 And _mask <= 32, "/" & _mask.ToString, String.Empty))
        End Function

        Public Shared Function TryParse(StringAddress As String, ByRef Address As IPAddress) As Boolean
            If String.IsNullOrEmpty(StringAddress) Then Return False
            Dim m As System.Text.RegularExpressions.Match = _
                System.Text.RegularExpressions.Regex.Match(StringAddress, "^(\d{1,3})\.(\d{1,3})\.(\d{1,3})\.(\d{1,3})(\/\d{1,2})?$")
            If m.Success = False Then Return False
            If Integer.Parse(m.Groups(1).Value) > 254 Then Return False
            If Integer.Parse(m.Groups(2).Value) > 254 Then Return False
            If Integer.Parse(m.Groups(3).Value) > 254 Then Return False
            If Integer.Parse(m.Groups(4).Value) > 254 Then Return False
            If m.Groups(5).Length > 0 AndAlso Integer.Parse(m.Groups(5).Value.Substring(1)) > 32 Then Return False
            Address = New IPAddress() With {._oct0 = Byte.Parse(m.Groups(1).Value), ._oct1 = Byte.Parse(m.Groups(2).Value), _
                                            ._oct2 = Byte.Parse(m.Groups(3).Value), ._oct3 = Byte.Parse(m.Groups(4).Value), _
                                            ._mask = 24}
            If m.Groups(5).Length > 0 Then Address._mask = Byte.Parse(m.Groups(5).Value.Substring(1))
            Return True
        End Function
        Public Shared Function TryParse(StringAddress As String, StringMask As String, ByRef Address As IPAddress) As Boolean
            If String.IsNullOrEmpty(StringMask) Then Return TryParse(StringAddress, Address)
            If Not TryParse(StringAddress, Address) Then Return False

            Dim m As System.Text.RegularExpressions.Match = _
                System.Text.RegularExpressions.Regex.Match(StringMask, "^(\d{1,3})\.(\d{1,3})\.(\d{1,3})\.(\d{1,3})$")
            If m.Success = False Then Return False
            If Integer.Parse(m.Groups(1).Value) > 255 Then Return False
            If Integer.Parse(m.Groups(2).Value) > 255 Then Return False
            If Integer.Parse(m.Groups(3).Value) > 255 Then Return False
            If Integer.Parse(m.Groups(4).Value) > 255 Then Return False

            Dim cidr As Integer = CInt(CInt(m.Groups(1).Value) << 24)
            cidr += CInt(CInt(m.Groups(2).Value) << 16)
            cidr += CInt(CInt(m.Groups(3).Value) << 8)
            cidr += CInt(m.Groups(4).Value)

            Dim bitcount As Integer = 0
            While cidr Mod 2 = 0
                bitcount += 1
                cidr = CInt(cidr >> 1)
            End While
            cidr = 32 - bitcount
            cidr = 32 - bitcount
            Address._mask = CByte(cidr)
            Return True
        End Function

    End Structure

End Class
