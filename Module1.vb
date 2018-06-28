Imports System.Management

Module Module1
    Sub Main()
        Dim cc As ConsoleColor = Console.ForegroundColor
        Console.ForegroundColor = ConsoleColor.DarkGray
        Console.WriteLine("Velwetowl Std Network management utility")
        Console.ForegroundColor = cc
        Dim wmi As New wmiObject
        If wmi.IsValid = False Then
            Console.WriteLine("Error: WMI subsystem not available. Press any key to exit...")
            Console.ReadKey()
        Else
            MainMenu(wmi)
        End If
        wmi.Dispose()
    End Sub
    Private Sub MainMenu(WMI As wmiObject)
        Console.WriteLine("Interfaces >>" & Environment.NewLine)
        Dim nicInd As Integer = 0
        For Each nic In WMI.NICs
            Console.WriteLine(("[" & nicInd.ToString & "]").PadRight(5) & nic.Name)
            nicInd += 1
        Next
        If nicInd = 0 Then
            Console.WriteLine("Network interfaces not found. Press any key to exit...")
            Console.ReadKey()
            Exit Sub
        End If

        Console.WriteLine(("[" & nicInd.ToString & "]").PadRight(5) & "Exit")
        Console.Write(Environment.NewLine & "Type option number and press [Enter]... ")
        Dim strInput As String, iInput As Byte = 255

        While True
            strInput = Console.ReadLine()
            If String.IsNullOrEmpty(strInput) Then Exit While
            If Byte.TryParse(strInput, iInput) = True AndAlso (iInput <= nicInd) Then Exit While
            iInput = 255
            Console.Write("Wrong input. Type option number and press [Enter]... ")
        End While
        strInput = Nothing
        If iInput = nicInd Or iInput = 255 Then Exit Sub
        NICMenu(WMI, WMI.NICs(iInput))
    End Sub

    Private Sub NICMenu(WMI As wmiObject, Nic As wmiNICObject)
        Try : Console.Clear() : Catch : End Try
        Console.WriteLine("Interfaces >> " & Nic.Name.Trim.ToUpper & " >> " & Environment.NewLine)
        Console.WriteLine("[0]  IP Address(-es):".PadRight(24) & Nic.IP)
        Console.WriteLine("[1]  Default gateway:".PadRight(24) & Nic.Gateway)
        Console.WriteLine("[2]  DNS Server(-s):".PadRight(24) & Nic.DNS)
        Console.WriteLine("[3]  WINS Server(-s):".PadRight(24) & Nic.WINS)
        Console.WriteLine("[4]  Return")
        Console.Write(Environment.NewLine & "Type option number and press [Enter]... ")

        Dim strInput As String, iInput As Byte = 255
        While True
            strInput = Console.ReadLine
            If String.IsNullOrEmpty(strInput) Then Exit While
            If Byte.TryParse(strInput, iInput) = True AndAlso (iInput >= 0 And iInput <= 4) Then Exit While
            iInput = 255
            Console.Write("Wrong input. Type option number and press [Enter]... ")
        End While
        strInput = Nothing
        If iInput = 255 Or iInput = 4 Then
            Try : Console.Clear() : Catch : End Try
            MainMenu(WMI)
        Else
            NICPropMenu(WMI, Nic, iInput)
        End If
    End Sub
    Private Sub NICPropMenu(WMI As wmiObject, Nic As wmiNICObject, PropIndex As Byte)
        Try : Console.Clear() : Catch : End Try
        Console.Write("Interfaces >> " & Nic.Name.Trim.ToUpper & " >> ")
        Select Case PropIndex
            Case 0
                Console.WriteLine("IP Address(-es)" & Environment.NewLine)
                Console.WriteLine("Current value: " & Nic.IP)
                Console.WriteLine("Type new IP and subnet mask (format xxx.xxx.xxx.xxx/xx, may specify multiple values delimited by comma) and press [Enter]" & _
                                  Environment.NewLine & "You may also type 'dhcp' or 'dynamic' to enable DHCP on adapter" & _
                                  Environment.NewLine & "Press just [Enter] to make no changes")
                Console.Write("> ")
            Case 1
                Console.WriteLine("Default Gateway" & Environment.NewLine)
                Console.WriteLine("Current value: " & Nic.Gateway)
                Console.WriteLine("Type new gateway address (format xxx.xxx.xxx.xxx) and press [Enter]" & _
                                  Environment.NewLine & "Press just [Enter] to make no changes")
                Console.Write("> ")
            Case 2
                Console.WriteLine("DNS Server(-s)" & Environment.NewLine)
                Console.WriteLine("Current value: " & Nic.DNS)
                Console.WriteLine("Type new DNS address list (format xxx.xxx.xxx.xxx, may specify multiple values delimited by comma) and press [Enter]" & _
                                  Environment.NewLine & "Press just [Enter] to make no changes")
                Console.Write("> ")
            Case 3
                Console.WriteLine("WINS Server(-s)" & Environment.NewLine)
                Console.WriteLine("Current value: " & Nic.WINS)
                Console.WriteLine("Type new WINS address list (format xxx.xxx.xxx.xxx, may specify one or two values delimited by comma) and press [Enter]" & _
                                  Environment.NewLine & "Press just [Enter] to make no changes")
                Console.Write("> ")
        End Select

        Dim strInput As String = Console.ReadLine
        If String.IsNullOrEmpty(strInput) Then
            NICMenu(WMI, Nic)
        Else
            Select Case PropIndex
                Case 0
                    Nic.IP = strInput.Trim
                Case 1
                    Nic.Gateway = strInput.Trim
                Case 2
                    Nic.DNS = strInput.Trim
                Case 3
                    Nic.WINS = strInput.Trim
            End Select
            Dim current_ind As Integer = WMI.NICs.IndexOf(Nic)
            WMI.Reset()
            If WMI.IsValid = False OrElse current_ind >= WMI.NICs.Count Then
                Console.Write("Error: WMI subsystem out of order. Press any key to exit...")
                Console.ReadKey()
            Else
                NICMenu(WMI, WMI.NICs(current_ind))
            End If
        End If
    End Sub

    'Private Function GetMgmtObj(MgmtClass As ManagementClass, ID As String) As ManagementObject
    '    For Each objMO As ManagementObject In MgmtClass.GetInstances
    '        If objMO("Caption").Equals(ID) Then Return objMO
    '    Next
    '    Return Nothing
    'End Function



    'Private Function GetMOPropertyValue(Collection As ManagementObjectCollection, ID As String, Prop As String) As Object
    '    Return GetMOPropertyValue(GetMO(Collection, ID), Prop)
    'End Function
    'Private Function GetMOPropertyValue(MO As ManagementObject, Prop As String) As Object
    '    If MO Is Nothing Then Return Nothing
    '    For Each objMOProp In MO.Properties
    '        If objMOProp.Name.Equals(Prop) Then Return objMOProp.Value
    '    Next
    '    Return Nothing
    'End Function


End Module


