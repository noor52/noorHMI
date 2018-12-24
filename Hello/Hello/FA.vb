

Imports System.Net
Imports System.Net.Sockets
Imports System.Text
Imports System.IO



Public Class FA
    Public Const STX = Chr(2)
    Public Const ETX = Chr(3)
    Public Const ENQ = Chr(5)
    ' Public gSocket As New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)
    '  Public gSocketWrite As New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)
    Public M(2000, 5) As Integer
    Public Y(2000, 5) As Integer
    Public X(2000, 5) As Integer
    Public R(2000, 5) As Integer
    Public D(2000, 5) As Integer
    Public activeListStations1() As String = New [String](2000) {}
    Public activeListStations2() As String = New [String](2000) {}
    Public activeListStations3() As String = New [String](2000) {}
    Public activeListStations4() As String = New [String](2000) {}
    Public activeListStations5() As String = New [String](2000) {}
    Public activeCountStations() As String = New [String](5) {}
    Public bufferedActiveCountStations() As String = New [String](5) {}
    Public bufferedSendSring() As String = New [String](20) {}
    Public PLCAddress(2000) As String
    Public IP As String
    Public AddressCount, IPport As Integer
    Public PLCScan As New Threading.Thread(AddressOf Engine)
    Public IPPing As New Threading.Thread(AddressOf Ping)
    Public RequestCount As Integer = 0
    ' Public WithEvents Scanner As New Timer
    Public Event R_Change()
    Public Event Connected()
    Public Event RX_Complete()
    Public Event Disconnected()
    Public Event Done()
    Public ConnectionStatus As Boolean = False

    Public Sub SetIp(ByVal IP_ As String)
        IP = IP_
        IPport = 500
    End Sub
    Public Sub Connect(ByVal IP_ As String)

        SetIp(IP_)
        'AddHandler Disconnected, AddressOf reConnectServer
        PLCScan.Start()

    End Sub
    Public Function checkExist(ByVal activeList() As String, ByVal entry As String, ByVal length As Integer) As Boolean
        For i As Integer = 0 To length
            If activeList(i) = entry Then
                Return True
            End If
        Next
        Return False
    End Function
    Public Sub sendInfo(ByVal type As String, ByVal plc_address As Integer, ByVal station As Integer)
        Dim count As String() = activeCountStations
        Dim newEntry As String = type & plc_address
        Select Case station
            Case 1
                If Not checkExist(activeListStations1, newEntry, count(station) - 1) Then
                    activeListStations1(count(station)) = newEntry
                    activeCountStations(station) += 1
                End If
            Case 2
                If Not checkExist(activeListStations2, newEntry, count(station) - 1) Then
                    activeListStations2(count(station)) = newEntry
                    activeCountStations(station) += 1
                End If
            Case 3
                If Not checkExist(activeListStations3, newEntry, count(station) - 1) Then
                    activeListStations3(count(station)) = newEntry
                    activeCountStations(station) += 1
                End If
            Case 4
                If Not checkExist(activeListStations4, newEntry, count(station) - 1) Then
                    activeListStations4(count(station)) = newEntry
                    activeCountStations(station) += 1
                End If
            Case 5
                If Not checkExist(activeListStations5, newEntry, count(station) - 1) Then
                    activeListStations5(count(station)) = newEntry
                    activeCountStations(station) += 1
                End If
        End Select
    End Sub
    Public Sub UpdateControls()
        For stn As Integer = 1 To 5
            If activeCountStations(stn) > 0 Then
                Select Case stn
                    Case 1 : updateControlsByStation(activeCountStations(stn), stn, activeListStations1)
                    Case 2 : updateControlsByStation(activeCountStations(stn), stn, activeListStations2)
                    Case 3 : updateControlsByStation(activeCountStations(stn), stn, activeListStations3)
                    Case 4 : updateControlsByStation(activeCountStations(stn), stn, activeListStations4)
                    Case 5 : updateControlsByStation(activeCountStations(stn), stn, activeListStations5)
                End Select

            End If
        Next
    End Sub
    Public Sub updateControlsByStation(ByVal ActiveCount As Integer, ByVal stationNO As Integer, ByVal activeList() As String)

        If ActiveCount <> bufferedActiveCountStations(stationNO) Then
            updateBufferString(ActiveCount, stationNO, activeList)
        End If
        Dim list() As String = New [String](63) {}
        Dim index As Integer = 0

        For i As Integer = 0 To ActiveCount - 1 Step 64
            Dim n As Integer = 0
            If ActiveCount - i > 63 Then
                n = 63
            Else
                n = ActiveCount - i - 1
            End If
            For j As Integer = 0 To n
                list(j) = activeList(i + j)
            Next
            'Dim s As String = Fatek_Read_Mixed(stationNO, list, n + 1)
            Dim s As String = bufferedSendSring(index)
            index += 1
            Dim vServerResponse As Byte() = getVServerResponse(s, IPport, IP)
            If Not vServerResponse Is Nothing Then
                FatekMixedDecode(vServerResponse, list, n + 1)
            End If
        Next

    End Sub
    Public Sub updateBufferString(ByVal ActiveCount As Integer, ByVal stationNO As Integer, ByVal activeList() As String)
        Dim list() As String = New [String](63) {}
        Dim index As Integer = 0

        For i As Integer = 0 To ActiveCount - 1 Step 64
            Dim n As Integer = 0
            If ActiveCount - i > 63 Then
                n = 63
            Else
                n = ActiveCount - i - 1
            End If
            For j As Integer = 0 To n
                list(j) = activeList(i + j)
            Next
            Dim s As String = Fatek_Read_Mixed(stationNO, list, n + 1)
            bufferedSendSring(index) = s
            index += 1
        Next
        bufferedActiveCountStations(stationNO) = ActiveCount
    End Sub
    Public Function getVServerResponse(ByVal pSendStr As String, ByVal pPort As Integer, ByVal IP As String) As Byte()
        Dim vServerResponse As Byte() = ConnectSendReciveStatic(Encoding.UTF8.GetBytes(pSendStr), pPort, IP)
        Dim c As Integer = 0
        While vServerResponse Is Nothing
            vServerResponse = ConnectSendReciveStatic(Encoding.UTF8.GetBytes(pSendStr), pPort, IP)
            c = c + 1
            If c > 2 Then
                ' MsgBox("Connection failed")

                If ConnectionStatus = True Then
                    ConnectionStatus = False
                    RaiseEvent Disconnected()
                End If
                Return Nothing
            End If
        End While

        If ConnectionStatus = False Then
            ConnectionStatus = True
            RaiseEvent Connected()
        End If

        Return vServerResponse
    End Function
    Public Function ReadDoubleRegister(ByVal address As Integer, ByVal station As Integer) As Single
        Try
            Dim x, y, z As Integer
            x = R(address, station)
            y = R(address + 1, station)
            z = (y * &H10000) + x
            If z = 0 Then Return 0


            Dim hexa As String = Hex(z)
            Dim raw As Byte() = New Byte(hexa.Length / 2 - 1) {}
            For i As Integer = 0 To raw.Length - 1
                raw(raw.Length - i - 1) = Convert.ToByte(hexa.Substring(i * 2, 2), 16)
            Next
            Dim f As Single = BitConverter.ToSingle(raw, 0)
            Return f
        Catch ex As Exception
            ' MsgBox(ex.Message)
            Return 0
        End Try



        Return 0
    End Function
    Public Sub Write32bit(ByVal Station As Integer, ByVal Value As Integer, ByVal Address As String)
        ' If Value > 65536 Then
        Dim DR As String

        Dim LB As String
        Dim HB As String
        ' Dim c As Integer

        Dim L As Integer
        Dim H As Integer
        Dim LAddress As String
        Dim HAddress As String
        LAddress = Address.ToString
        HAddress = (Address + 1).ToString

        DR = Hex(Value).ToString
        DR = EightFormate(DR)
        HB = DR.Substring(0, 4)
        LB = DR.Substring(4, 4)

        L = HEX_to_DEC(LB)
        H = HEX_to_DEC(HB)


        SetSingleItem(Station, "R", LAddress, L)
        SetSingleItem(Station, "R", HAddress, H)


    End Sub
    Public Function Read32bit(ByVal Station As Integer, ByVal Address As Integer) As Int32
        Dim LB As String = Hex(R(Address, Station)).ToString
        Dim HB As String = Hex(R(Address + 1, Station)).ToString
        Dim val As String = UCase(HB & LB)
        Dim B As Int32 = 0
        For i = 1 To Len(val)
            Select Case Mid(val, Len(val) - i + 1, 1)
                Case "0" : B = B + 16 ^ (i - 1) * 0
                Case "1" : B = B + 16 ^ (i - 1) * 1
                Case "2" : B = B + 16 ^ (i - 1) * 2
                Case "3" : B = B + 16 ^ (i - 1) * 3
                Case "4" : B = B + 16 ^ (i - 1) * 4
                Case "5" : B = B + 16 ^ (i - 1) * 5
                Case "6" : B = B + 16 ^ (i - 1) * 6
                Case "7" : B = B + 16 ^ (i - 1) * 7
                Case "8" : B = B + 16 ^ (i - 1) * 8
                Case "9" : B = B + 16 ^ (i - 1) * 9
                Case "A" : B = B + 16 ^ (i - 1) * 10
                Case "B" : B = B + 16 ^ (i - 1) * 11
                Case "C" : B = B + 16 ^ (i - 1) * 12
                Case "D" : B = B + 16 ^ (i - 1) * 13
                Case "E" : B = B + 16 ^ (i - 1) * 14
                Case "F" : B = B + 16 ^ (i - 1) * 15
            End Select
        Next i
        Return B
    End Function

    Public Function EightFormate(ByVal STR As String) As String
        Dim a As String = ""
        Dim i As Integer

        i = Len(STR)

        Select Case i

            Case 1 : a = "0000000" & STR
            Case 2 : a = "000000" & STR
            Case 3 : a = "00000" & STR
            Case 4 : a = "0000" & STR
            Case 5 : a = "000" & STR
            Case 6 : a = "00" & STR
            Case 7 : a = "0" & STR
            Case 8 : a = "" & STR



        End Select


        Return a

    End Function
    Public Sub Disconnect()
        PLCScan.Abort()

    End Sub

    Public Function GetRXData(ByVal Device As String) As Array
        Select Case Device
            Case "R"
                Return R

            Case "M"
                Return M

            Case "X"
                Return X
            Case "Y"
                Return Y
        End Select
        Return Nothing

    End Function

    Public Sub Ping()
        Do
            Dim p As New System.Net.NetworkInformation.Ping()

            Dim reply As System.Net.NetworkInformation.PingReply = p.Send(IP)


            If reply.Status = System.Net.NetworkInformation.IPStatus.Success Then

            Else

            End If

        Loop
    End Sub
    Public Sub AddAddress(ByVal AddressType As String, ByVal Address As Integer, ByVal length As Integer, ByVal Station As String)
        PLCAddress(AddressCount) = AddressType & "," & Address & "," & length & "," & Station
        AddressCount += 1
    End Sub

    Public Sub AddressProcess(ByVal Address As String)
        Dim AT As String
        Dim A As Integer
        Dim AL As Integer
        Dim ST As String

        AT = GetAddressType(Address)
        A = GetAddress(Address)
        AL = GetAddressLenth(Address)
        ST = GetStation(Address)
        GetItemBlock(AT, A, AL, IP, ST)

    End Sub

    Public Function GetAddressType(ByVal STR As String)

        Dim arr As String() = STR.Split(",")

        If arr.Length = 0 Then
            Return Nothing
            Exit Function

        End If

        Return arr(0)

    End Function

    Public Function GetAddress(ByVal STR As String) As String

        Dim arr As String() = STR.Split(",")

        If arr.Length = 0 Then
            Return Nothing
            Exit Function

        End If

        Return arr(1)

    End Function

    Public Function GetAddressLenth(ByVal STR As String) As String

        Dim arr As String() = STR.Split(",")

        If arr.Length = 0 Then

            Return Nothing
            Exit Function

        End If

        Return arr(2)

    End Function
    Public Function GetStation(ByVal STR As String)

        Dim arr As String() = STR.Split(",")

        If arr.Length = 0 Then
            Return Nothing

            Exit Function

        End If

        Return arr(3)

    End Function


    Public Sub Engine()

        'Dim i As Integer

        Do
            UpdateControls()
            'For i = 0 To AddressCount - 1
            '    AddressProcess(PLCAddress(i))
            'Next
        Loop
    End Sub


    Public Sub GetItemBlock(ByVal addressType As String, ByVal Address As Integer, ByVal length As String, ByVal IP As String, ByVal Station As String) 'As Integer()
        Dim a As String
        a = Fatek_Read_PLC_Message(Station, addressType, Address, length)
        Dim vServerResponse As Byte() = ConnectSendReciveStatic(Encoding.UTF8.GetBytes(a), 500, IP)
        'If vServerResponse Is Nothing Then
        '    RaiseEvent Disconnected()

        '    Exit Sub
        'End If
        Dim i As Integer = 0
        While vServerResponse Is Nothing
            vServerResponse = ConnectSendReciveStatic(Encoding.UTF8.GetBytes(a), 500, IP)
            i = i + 1
            If i > 2 Then
                If ConnectionStatus = True Then
                    ConnectionStatus = False
                    RaiseEvent Disconnected()
                End If
                Exit Sub
            End If
        End While
        'MsgBox(i)

        If ConnectionStatus = False Then
            ConnectionStatus = True
            RaiseEvent Connected()
        End If
        If Recive_SUMCHECK(vServerResponse, addressType) = True Then
            FatekBlockDecode(addressType, vServerResponse, Address, 64, Station)
            RaiseEvent RX_Complete()
        End If

    End Sub
    Public Function Recive_SUMCHECK(ByVal D As Byte(), ByVal AD_TYPE As String) As Boolean
        Dim I, length As Integer
        length = 0
        Dim BCC, BCCC, RX_ADDRESSTYPR As String
        BCC = ""
        Dim S, S2 As String
        S2 = ""
        S = ""
        RX_ADDRESSTYPR = Chr(D(3))
        RX_ADDRESSTYPR = RX_ADDRESSTYPR & Chr(D(4))

        For I = 0 To D.Length - 1
            S2 = S2 & Chr(D(I))
        Next

        For I = 0 To D.Length - 4
            S = S & Chr(D(I))
        Next
        For I = D.Length - 3 To D.Length - 2
            BCC = BCC & Chr(D(I))
        Next
        BCCC = SumChk(S)
        If BCC = BCCC Then
            Return True
        End If
        Return False
    End Function


    Public Function HEX_to_DEC(ByVal Hex_Renamed As String) As Integer '16#to10#
        Dim i As Integer
        Dim B As Integer
        Hex_Renamed = UCase(Hex_Renamed)
        For i = 1 To Len(Hex_Renamed)
            Select Case Mid(Hex_Renamed, Len(Hex_Renamed) - i + 1, 1)
                Case "0" : B = B + 16 ^ (i - 1) * 0
                Case "1" : B = B + 16 ^ (i - 1) * 1
                Case "2" : B = B + 16 ^ (i - 1) * 2
                Case "3" : B = B + 16 ^ (i - 1) * 3
                Case "4" : B = B + 16 ^ (i - 1) * 4
                Case "5" : B = B + 16 ^ (i - 1) * 5
                Case "6" : B = B + 16 ^ (i - 1) * 6
                Case "7" : B = B + 16 ^ (i - 1) * 7
                Case "8" : B = B + 16 ^ (i - 1) * 8
                Case "9" : B = B + 16 ^ (i - 1) * 9
                Case "A" : B = B + 16 ^ (i - 1) * 10
                Case "B" : B = B + 16 ^ (i - 1) * 11
                Case "C" : B = B + 16 ^ (i - 1) * 12
                Case "D" : B = B + 16 ^ (i - 1) * 13
                Case "E" : B = B + 16 ^ (i - 1) * 14
                Case "F" : B = B + 16 ^ (i - 1) * 15
            End Select
        Next i
        HEX_to_DEC = B
    End Function

    Public Function FatekBlockDecode(ByVal type As String, ByVal ReceivedData As Byte(), ByVal index As Integer, ByVal length As Integer, ByVal Station As String)
        Dim i As Integer
        Dim s As String

        Dim readMode As String = Chr(ReceivedData(3)) & Chr(ReceivedData(4))
        Dim RX_STATION As String = Chr(ReceivedData(1)) & Chr(ReceivedData(2))


        Select Case type

            Case "R"
                If readMode = "46" Then
                    For i = 6 To ReceivedData.Length - 4 Step 4
                        s = Chr(ReceivedData(i)) & Chr(ReceivedData(i + 1)) & Chr(ReceivedData(i + 2)) & Chr(ReceivedData(i + 3))

                        R(index, RX_STATION) = HEX_to_DEC(s)
                        index += 1
                    Next
                End If
            Case "D"
                If readMode = "46" Then
                    For i = 6 To ReceivedData.Length - 4 Step 4
                        s = Chr(ReceivedData(i)) & Chr(ReceivedData(i + 1)) & Chr(ReceivedData(i + 2)) & Chr(ReceivedData(i + 3))

                        D(index, RX_STATION) = HEX_to_DEC(s)
                        index += 1
                    Next
                End If
            Case "M"
                If readMode = "44" Then
                    For i = 6 To ReceivedData.Length - 4
                        s = Chr(ReceivedData(i))

                        M(index, RX_STATION) = HEX_to_DEC(s)

                        index += 1
                    Next
                End If
            Case "Y"
                If readMode = "44" Then
                    For i = 6 To ReceivedData.Length - 4
                        s = Chr(ReceivedData(i))

                        Y(index, RX_STATION) = HEX_to_DEC(s)

                        index += 1

                    Next
                End If

            Case "X"
                If readMode = "44" Then
                    For i = 6 To ReceivedData.Length - 4
                        s = Chr(ReceivedData(i))


                        X(index, RX_STATION) = s

                        index += 1
                    Next
                End If
        End Select

        Return Nothing
    End Function
    Public Function FatekMixedDecode(ByVal ReceivedData As Byte(), ByVal mixedAddressArray As String(), ByVal Length As Integer)
        Dim i, j As Integer
        Dim s As String
        Dim v As Integer

        Dim RX_STATION As String
        RX_STATION = Chr(ReceivedData(1)) & Chr(ReceivedData(2))

        i = 6
        j = 0

        While i <= ReceivedData.Length - 4 And j < Length
            Dim Type As String = mixedAddressArray(j).Substring(0, 1)
            Dim Address As String = mixedAddressArray(j).Substring(1)

            Select Case Type

                Case "R"
                    s = Chr(ReceivedData(i)) & Chr(ReceivedData(i + 1)) & Chr(ReceivedData(i + 2)) & Chr(ReceivedData(i + 3))

                    v = HEX_to_DEC(s)

                    R(Address, RX_STATION) = v
                    i = i + 4
                Case "D"
                    s = Chr(ReceivedData(i)) & Chr(ReceivedData(i + 1)) & Chr(ReceivedData(i + 2)) & Chr(ReceivedData(i + 3))

                    v = HEX_to_DEC(s)
                    i = i + 4
                    D(Address, RX_STATION) = v
                Case "M"
                    s = Chr(ReceivedData(i))

                    v = HEX_to_DEC(s)
                    i = i + 1
                    M(Address, RX_STATION) = v
                Case "Y"
                    s = Chr(ReceivedData(i))

                    v = HEX_to_DEC(s)

                    Y(Address, RX_STATION) = v
                    i = i + 1

                Case "X"
                    s = Chr(ReceivedData(i))

                    v = HEX_to_DEC(s)
                    i = i + 1
                    X(Address, RX_STATION) = v
            End Select
            j = j + 1
        End While


        Return Nothing
    End Function


    Public Function ConnectSendRecive(ByVal pSendData As Byte(), ByVal pIp As String, ByVal pPort As Integer, ByVal IP As String) As Byte()
        If String.IsNullOrEmpty(pIp) Then Return Nothing
        If Not IPAddress.TryParse(pIp, Nothing) Then Return Nothing
        Dim vIp As IPAddress = IPAddress.Parse(IP)
        Dim vEndPoint As New IPEndPoint(vIp, CInt("500"))
        Dim vTimeout As Integer = 500
        Dim vMessageLenth As Integer = 1024


        Dim vServerResponse As Byte() = Nothing

        Dim gSocket = New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)



        If Not gSocket.Connected Then
            Try
                gSocket.SendTimeout = vTimeout
                gSocket.ReceiveTimeout = vTimeout
                gSocket.Connect(vEndPoint)
            Catch ex As Exception
                If ConnectionStatus = True Then
                    ConnectionStatus = False
                    RaiseEvent Disconnected()
                End If
                Return Nothing
            End Try
        End If
        ' End While

        If Not gSocket.Connected Then Return Nothing
        'send

        gSocket.Send(pSendData)


        'recive
        '  log("Send Request " & RequestCount)
        Dim vBuffer(vMessageLenth - 1) As Byte
        Dim vNumOfBytesRecived As Integer = 0

        Try
            vNumOfBytesRecived = gSocket.Receive(vBuffer, 0, vMessageLenth, SocketFlags.None)

        Catch ex As Exception
            Return Nothing
        End Try
        ' log("Rcv Request " & RequestCount)

        '  Return recived data
        ReDim vServerResponse(vNumOfBytesRecived - 1)
        Array.Copy(vBuffer, vServerResponse, vNumOfBytesRecived)

        gSocket.Shutdown(SocketShutdown.Both)
        gSocket.Close()
        ' gSocket.Disconnect(True)
        '    End Using
        ' End If
        Return vServerResponse
    End Function

    Public Function ConnectSendReciveWrite(ByVal pSendData As Byte(), ByVal pIp As String, ByVal pPort As Integer, ByVal IP As String) As Byte()
        If String.IsNullOrEmpty(pIp) Then Return Nothing
        If Not IPAddress.TryParse(pIp, Nothing) Then Return Nothing
        Dim vIp As IPAddress = IPAddress.Parse(IP)
        Dim vEndPoint As New IPEndPoint(vIp, CInt("500"))
        Dim vTimeout As Integer = 500
        Dim vMessageLenth As Integer = 1024


        Dim vServerResponse As Byte() = Nothing

        Dim gSocketWrite As New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)



        If Not gSocketWrite.Connected Then
            Try
                gSocketWrite.SendTimeout = vTimeout
                gSocketWrite.ReceiveTimeout = vTimeout
                gSocketWrite.Connect(vEndPoint)
            Catch ex As Exception
                If ConnectionStatus = True Then
                    ConnectionStatus = False
                    RaiseEvent Disconnected()
                End If
                Return Nothing
            End Try
        End If
        ' End While

        If Not gSocketWrite.Connected Then Return Nothing
        'send

        gSocketWrite.Send(pSendData)

        'recive
        '  log("Send Request " & RequestCount)
        Dim vBuffer(vMessageLenth - 1) As Byte
        Dim vNumOfBytesRecived As Integer = 0

        Try
            vNumOfBytesRecived = gSocketWrite.Receive(vBuffer, 0, vMessageLenth, SocketFlags.None)

        Catch ex As Exception
            Return Nothing
        End Try
        '  log("Rcv Request " & RequestCount)

        '  Return recived data
        ReDim vServerResponse(vNumOfBytesRecived - 1)
        Array.Copy(vBuffer, vServerResponse, vNumOfBytesRecived)

        gSocketWrite.Shutdown(SocketShutdown.Both)
        gSocketWrite.Close()
        ' gSocket.Disconnect(True)
        '    End Using
        ' End If
        Return vServerResponse
    End Function


    Private Sub log(ByVal txt As String)
        'Dim strFile As String = "E:\ErrorLog_" & DateTime.Today.ToString("dd-MMM-yyyy") & ".txt"

        'Dim sw As StreamWriter
        'Dim fs As FileStream = Nothing
        'Dim s As New Object
        'If (Not File.Exists(strFile)) Then
        '    Try
        '        fs = File.Create(strFile)
        '        sw = File.AppendText(strFile)

        '        sw.WriteLine(txt & ": " & DateTime.Now)
        '    Catch ex As Exception
        '        MsgBox(ex.ToString)
        '    End Try
        'Else

        '    SyncLock s
        '        Try
        '            sw = File.AppendText(strFile)
        '            sw.WriteLine(txt & ": " & DateTime.Now)

        '            sw.Close()
        '        Catch ex As Exception
        '            MsgBox(ex.ToString)
        '        End Try
        '    End SyncLock
        'End If
    End Sub




    Public Function Fatek_Read_PLC_Message(ByVal Station As String, ByVal DeviceType As String, ByVal Address As String, ByVal length As String) As String
        Dim S As String = ""
        Dim sum As String = ""
        Dim C As String = ""

        Select Case DeviceType
            Case "R"
                S = STX & FormatString(Station, 2) & "46" & FormatString(Hex(length), 2) & DeviceType & FormatString(Address, 5)
            Case "M"
                S = STX & FormatString(Station, 2) & "44" & FormatString(Hex(length), 2) & DeviceType & FormatString(Address, 4)
            Case "Y"
                S = STX & FormatString(Station, 2) & "44" & FormatString(Hex(length), 2) & DeviceType & FormatString(Address, 4)
            Case "X"
                S = STX & FormatString(Station, 2) & "44" & FormatString(Hex(length), 2) & DeviceType & FormatString(Address, 4)
        End Select

        sum = SumChk(S)
        Return S & sum & ETX
    End Function


    Public Function Fatek_Read_Mixed(ByVal Station As String, ByVal mixedAddressArray As String(), ByVal length As String) As String
        Dim S, Address, DeviceType As String
        Dim sum As String = ""
        Dim i As Integer = 0
        S = STX & FormatString(Station, 2) & "48" & FormatString(Hex(length), 2)

        For i = 0 To length - 1
            DeviceType = mixedAddressArray(i).Substring(0, 1)
            Address = mixedAddressArray(i).Substring(1)
            Select Case DeviceType
                Case "R"
                    S = S & DeviceType & FormatString(Address, 5)
                Case "M"
                    S = S & DeviceType & FormatString(Address, 4)
                Case "Y"
                    S = S & DeviceType & FormatString(Address, 4)
                Case "X"
                    S = S & DeviceType & FormatString(Address, 4)
            End Select
        Next
        sum = SumChk(S)
        Return S & sum & ETX
    End Function

    Public Function Fatek_Write_PLC_Message_Single(ByVal Station As String, ByVal DeviceType As String, ByVal Address As String, ByVal length As String, ByVal data As Integer) As String
        Dim S As String = ""
        Dim sum As String

        Select Case DeviceType
            Case "R"
                S = STX & FormatString(Station, 2) & "47" & FormatString(length, 2) & DeviceType & FormatString(Address, 5) & FormatString(Hex(data), 4)
            Case "M"
                S = STX & FormatString(Station, 2) & "45" & "01" & DeviceType & FormatString(Address, 4) & FormatString(Hex(data), 1)
            Case "Y"
                S = STX & FormatString(Station, 2) & "45" & "01" & DeviceType & FormatString(Address, 4) & FormatString(Hex(data), 1)
        End Select
        sum = SumChk(S)
        Return S & sum & ETX

    End Function

    Public Function Fatek_Write_PLC_Message_Multi(ByVal Station As String, ByVal DeviceType As String, ByVal Address As String, ByVal length As String, ByVal data() As Integer) As String
        Dim S As String = ""
        Dim sum As String = ""
        Dim dataString As String

        dataString = DataFormat(data)
        Select Case DeviceType
            Case "R"
                S = STX & FormatString(Station, 2) & "47" & FormatString(length, 2) & DeviceType & FormatString(Address, 5) & dataString ' FormatString(Hex(data), 4)
            Case "M"
                S = STX & FormatString(Station, 2) & "45" & "01" & DeviceType & FormatString(Address, 4) & FormatString(Hex(data), 1)
            Case "Y"
                S = STX & FormatString(Station, 2) & "45" & "01" & DeviceType & FormatString(Address, 4) & FormatString(Hex(data), 1)
        End Select
        sum = SumChk(S)
        Return S & sum & ETX

    End Function

    Public Function FormatString(ByVal Value As String, ByVal length As Integer) As String

        If length = 4 Then
            Select Case Len(Value)
                Case 1 : Value = "000" & Value
                Case 2 : Value = "00" & Value
                Case 3 : Value = "0" & Value
                Case 4 : Value = "" & Value
            End Select
            Return Value
        End If

        If length = 5 Then
            Select Case Len(Value)
                Case 1 : Value = "0000" & Value
                Case 2 : Value = "000" & Value
                Case 3 : Value = "00" & Value
                Case 4 : Value = "0" & Value
                Case 5 : Value = "" & Value
            End Select
            Return Value
        End If

        If length = 2 Then
            Select Case Len(Value)

                Case 1 : Value = "0" & Value
                Case 2 : Value = "" & Value
            End Select

        End If

        Return Value
        Exit Function
    End Function


    Public Function DataFormat(ByVal data As Array) As String
        Dim a As String = ""
        Dim l As Integer = data.Length

        For i = 0 To l - 1
            a = a & FormatString(Hex(data(i)), 4)
        Next

        Return a
    End Function

    Public Function SumChk(ByRef Dats As String) As String 'Sum Check
        Dim i As Integer
        Dim CHK As Integer
        For i = 1 To Len(Dats)
            CHK = CHK + Asc(Mid(Dats, i, 1))
        Next i
        SumChk = Right(Hex(CHK), 2)
    End Function

    Public Function SetSingleItem(ByVal Station As Integer, ByVal addressType As String, ByVal Address As String, ByVal Data As Integer) As Boolean
        If ConnectionStatus = True Then
            Dim a As String = Fatek_Write_PLC_Message_Single(Station, addressType, Address, "01", Data)
            For i As Integer = 1 To 3
                If WriteData(a) = True Or GetSingleAcknowledge(Station, addressType, Address, Data) = True Then Return True
            Next
        Else
            MsgBox("Server Disconnected")
        End If
        Return False

    End Function
    Public Function GetMultipleRegisterAcknowledge(ByVal Address As String, ByVal Data() As Integer, ByVal Length As Integer) As Boolean
        For i As Integer = 0 To Length - 1
            If R(Address + i, 1) <> Data(i) Then Return False
        Next
        Return True
    End Function
    Public Function GetSingleAcknowledge(ByVal Station As Integer, ByVal addressType As String, ByVal Address As String, ByVal Data As Integer) As Boolean
        Select Case addressType
            Case "R"
                If R(Address, Station) = Data Then Return True
            Case "Y"
                If Y(Address, Station) = Data Then Return True
            Case "M"
                If M(Address, Station) = Data Then Return True
        End Select
        Return False
    End Function
    Public Function WriteData(ByVal a As String) As Boolean
        Dim vServerResponse As Byte() = ConnectSendReciveStatic(Encoding.UTF8.GetBytes(a), 500, IP)
        Dim i As Integer = 0
        While vServerResponse Is Nothing
            vServerResponse = ConnectSendReciveStatic(Encoding.UTF8.GetBytes(a), 500, IP)
            If vServerResponse Is Nothing Then
                i = i + 1
            End If
            If i > 2 Then
                If ConnectionStatus = True Then
                    ConnectionStatus = False
                    RaiseEvent Disconnected()
                End If
                Exit Function
            End If
        End While
        If ConnectionStatus = False Then
            ConnectionStatus = True
            RaiseEvent Connected()
        End If
        If vServerResponse.Length = 9 And vServerResponse(5) = 48 Then
            Return True
        End If

        Return False
    End Function
    Public Function SetMultyItem(ByVal addressType As String, ByVal Address As String, ByVal Data() As Integer, ByVal Length As Integer) As Boolean
        If ConnectionStatus = True Then

            Dim a As String = Fatek_Write_PLC_Message_Multi("1", addressType, Address, Length, Data)
            For i As Integer = 1 To 3
                If WriteData(a) = True Or GetMultipleRegisterAcknowledge(Address, Data, Length) = True Then Return True
            Next
        Else
            MsgBox("Server Disconnected")
        End If


        Return False
    End Function

    Public Sub reConnectServer()
        'Dim vIp As IPAddress = IPAddress.Parse(IP)
        'Dim vEndPoint As New IPEndPoint(vIp, CInt("500"))


        ' RequestCount += 1
        ' log(RequestCount)
        'If Not gSocket.Connected Then
        '    'gSocket.Close()
        '    gSocket = New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)
        '    Try
        '        gSocket.Connect(vEndPoint)
        '        '  log("recon...........................................................................................")
        '    Catch ex As Exception
        '        RaiseEvent Disconnected()
        '        ConnectionStatus = False
        '    End Try
        'Else
        '    RaiseEvent Connected()
        '    ConnectionStatus = True
        '    Exit Sub
        'End If
        'If Not gSocket.Connected Then
        '    Try
        '        gSocket.Connect(vEndPoint)
        '    Catch ex As Exception
        '        RaiseEvent Disconnected()
        '        ConnectionStatus = False

        '    End Try
        'Else
        '    RaiseEvent Connected()
        '    ConnectionStatus = True
        'End If
        'If Not gSocketWrite.Connected Then
        '    'gSocket.Close()
        '    gSocketWrite = New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)
        '    Try
        '        gSocketWrite.Connect(vEndPoint)
        '        ' log("recon...........................................................................................")
        '    Catch ex As Exception
        '        RaiseEvent Disconnected()
        '        ConnectionStatus = False
        '    End Try
        'Else
        '    RaiseEvent Connected()
        '    ConnectionStatus = True
        '    Exit Sub
        'End If
        'If Not gSocketWrite.Connected Then
        '    Try
        '        gSocketWrite.Connect(vEndPoint)
        '    Catch ex As Exception
        '        RaiseEvent Disconnected()
        '        ConnectionStatus = False

        '    End Try
        'Else
        '    RaiseEvent Connected()
        '    ConnectionStatus = True
        'End If
    End Sub
    Public Shared Function ConnectSendReciveStatic(ByVal pSendData As Byte(), ByVal pPort As Integer, ByVal IP As String) As Byte()

        Dim vIp As IPAddress = IPAddress.Parse(IP)
        Dim vEndPoint As New IPEndPoint(vIp, Convert.ToInt32("500"))
        Dim vTimeout As Integer = 500
        Dim vMessageLenth As Integer = 1024


        Dim vServerResponse As Byte() = Nothing

        Using gSocket As New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)

            Try

                gSocket.Connect(vEndPoint)
            Catch ex As Exception

                Return Nothing
            End Try


            If Not gSocket.Connected Then
                Return Nothing
            End If
            'send
            gSocket.SendTimeout = vTimeout
            gSocket.Send(pSendData)

            'recive
            gSocket.ReceiveTimeout = vTimeout
            Dim vBuffer As Byte() = New Byte(vMessageLenth - 1) {}
            Dim vNumOfBytesRecived As Integer = 0

            Try

                vNumOfBytesRecived = gSocket.Receive(vBuffer, 0, vMessageLenth, SocketFlags.None)
            Catch ex As Exception
                Return Nothing
            End Try

            vServerResponse = New Byte(vNumOfBytesRecived - 1) {}


            Array.Copy(vBuffer, vServerResponse, vNumOfBytesRecived)
        End Using

        Return vServerResponse
    End Function



    Public Function GetItem(ByVal a As String, ByVal Address As String) As Integer
        Dim adtype As String
        Dim addressint As Integer
        adtype = Address.Substring(0, 1)
        addressint = Address.Substring(1, (Len(Address)) - 1)

        Select Case adtype
            Case "R"
                Return R(addressint, 1)

            Case "M"
                Return M(addressint, 1)
            Case "X"
                Return X(addressint, 1)

            Case "Y"
                Return Y(addressint, 1)
        End Select



    End Function

    Public Sub Additem(ByVal a As String, ByVal Address As String)

        Dim arr As String() = Address.Split("-")

        If arr.Length = 1 Then
            Dim adtype As String
            Dim addressint As Integer
            adtype = Address.Substring(0, 1)
            addressint = Address.Substring(1, (Len(Address)) - 1)
            sendInfo(adtype, addressint, 1)
        ElseIf arr.Length = 2 Then
            Dim adtype1, adtype2 As String
            Dim addressint1, addressint2 As Integer
            adtype1 = arr(0).Substring(0, 1)
            adtype2 = arr(1).Substring(0, 1)
            addressint1 = arr(0).Substring(1, (Len(arr(0))) - 1)
            addressint2 = arr(1).Substring(1, (Len(arr(1))) - 1)
            If adtype1 = adtype2 Then





                For i = addressint1 To addressint2
                    sendInfo(adtype1, i, 1)
                Next


            End If

        End If


    End Sub

    Public Function SetItem(ByVal a As String, ByVal Address As String, ByVal val As Integer) As Integer
        Dim adtype As String
        Dim addressint As Integer
        adtype = Address.Substring(0, 1)
        addressint = Address.Substring(1, (Len(Address)) - 1)

        Select Case adtype
            Case "R"

                SetSingleItem(1, adtype, addressint, val)
            Case "M"
                SetSingleItem(1, adtype, addressint, val)
        End Select



    End Function





End Class
