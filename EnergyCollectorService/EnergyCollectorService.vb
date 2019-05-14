Imports System.Data.OleDb

Public Class EnergyCollectorService
    Dim GetDataIsProgress As Boolean
    Private eventId As Integer = 0
    Private ReadOnly timer As Timers.Timer = New Timers.Timer With {
            .Enabled = True,
            .Interval = 10 * 60000,
            .AutoReset = True
        }
    Protected Overrides Sub OnStart(ByVal args() As String)
        ' Добавьте здесь код запуска службы. Этот метод должен настроить все необходимое
        ' для успешной работы службы.
        GetDataIsProgress = False
        EventLog1.WriteEntry("Служба запущена", EventLogEntryType.Information, Math.Max(Threading.Interlocked.Increment(eventId), eventId - 1))
        'Создаем таймер 
        'Запускаем Ontimer при срабатывании таймера 
        AddHandler timer.Elapsed, AddressOf OnTimer

        timer.Start()
        '        GetData()
    End Sub

    Public Sub OnTimer(sender As Object, args As Timers.ElapsedEventArgs)
        EventLog1.WriteEntry(Now.ToString & " Запуск события по таймеру", EventLogEntryType.Information, Math.Max(Threading.Interlocked.Increment(eventId), eventId - 1))
        If Not GetDataIsProgress Then
            GetData()
        Else
            EventLog1.WriteEntry(Now.ToString & " Предыдущий опрос не окончен", EventLogEntryType.Information, Math.Max(Threading.Interlocked.Increment(eventId), eventId - 1))
            'GetDataIsProgress = False
        End If
    End Sub

    Protected Overrides Sub OnStop()
        timer.Stop()
        timer.Close()
        timer.Dispose()
        EventLog1.WriteEntry("Служба остановлена", EventLogEntryType.Information, Math.Max(Threading.Interlocked.Increment(eventId), eventId - 1))


    End Sub

    Public Sub GetData()

        Dim SerialNumber, CollectorType, BaudRate, SerialNumberFromDB, CollectorGroupFromDB, CollectorTypeFromDB, BaudRateFromDB As Integer
        Dim AnswerisOk As Boolean

        Dim ComPortName, Address, ComPortNameFromDB, AddressFromDB, DescriptionFromDB As String
        Dim Collectors As New List(Of Collector)()

        ' Энергия от сброса
        Dim AskSummaryEnergy As Byte() = {5, 0, 0}
        ' Энергия на начало года
        Dim AskYearEnergy As Byte() = {5, &H90, 0}
        ' Энергия на начало месяца
        Dim Month As Byte = &HB0 + Today.Month
        Dim AskMonthEnergy As Byte() = {5, Month, 0}
        ' Энергия на начало суток
        Dim AskDayEnergy As Byte() = {5, &HC0, 0}
        Dim dbConn As New OleDbConnection()

        GetDataIsProgress = True

        Try
            ' Подключаемся к базе
            dbConn.ConnectionString = "Provider=SQLOLEDB; Data Source=192.168.10.32\SQLEXPRESS; Initial Catalog=energy; Persist Security Info=False; User Id=energy; Password=energy"
            dbConn.Open()
        Catch ex As Exception
            EventLog1.WriteEntry("Ошибка соединения с базой данных: " + ex.Message, EventLogEntryType.Error, Math.Max(Threading.Interlocked.Increment(eventId), eventId - 1))
            dbConn.Close()
            GetDataIsProgress = False
            Exit Sub
        End Try

        Try
            ' Делаем запрос к базе 
            Dim dbCommandCollectors As OleDbCommand = dbConn.CreateCommand()
            dbCommandCollectors.CommandText = "SELECT * FROM Collectors"
            dbCommandCollectors.ExecuteNonQuery()
            Dim reader As OleDbDataReader = dbCommandCollectors.ExecuteReader()
            ' Читаем данные
            While (reader.Read())
                ' Если пустое значение, то присваиваем переменной 0
                If Not reader.IsDBNull(reader.GetOrdinal("SerialNumber")) Then
                    SerialNumberFromDB = reader(reader.GetOrdinal("SerialNumber"))
                Else
                    SerialNumberFromDB = 0
                End If

                If Not reader.IsDBNull(reader.GetOrdinal("ComPortName")) Then
                    ComPortNameFromDB = reader(reader.GetOrdinal("ComPortName"))
                Else
                    ComPortNameFromDB = ""
                End If

                If Not reader.IsDBNull(reader.GetOrdinal("Address")) Then
                    AddressFromDB = reader(reader.GetOrdinal("Address"))
                Else
                    AddressFromDB = ""
                End If

                If Not reader.IsDBNull(reader.GetOrdinal("CollectorGroup")) Then
                    CollectorGroupFromDB = reader(reader.GetOrdinal("CollectorGroup"))
                Else
                    CollectorGroupFromDB = 0
                End If

                If Not reader.IsDBNull(reader.GetOrdinal("Description")) Then
                    DescriptionFromDB = reader(reader.GetOrdinal("Description"))
                Else
                    DescriptionFromDB = ""
                End If

                If Not reader.IsDBNull(reader.GetOrdinal("BaudRate")) Then
                    BaudRateFromDB = reader(reader.GetOrdinal("BaudRate"))
                Else
                    BaudRateFromDB = 0
                End If

                If Not reader.IsDBNull(reader.GetOrdinal("CollectorType")) Then
                    CollectorTypeFromDB = reader(reader.GetOrdinal("CollectorType"))
                Else
                    CollectorTypeFromDB = 0
                End If

                EventLog1.WriteEntry(Now.ToString _
                                     & " Добавлен счетчик " _
                                     & " Компорт: " & ComPortNameFromDB & vbCrLf _
                                     & " Скорость компорта: " & BaudRateFromDB & vbCrLf _
                                     & " Тип счетчика: " & CollectorTypeFromDB & vbCrLf _
                                     & " Адрес счетчика: " & AddressFromDB & vbCrLf _
                                     & " Серийный номер: " & SerialNumberFromDB & vbCrLf _
                                     & " Группа: " & CollectorGroupFromDB & vbCrLf _
                                     & " Описание: " & DescriptionFromDB & vbCrLf _
                                     , EventLogEntryType.Information, Math.Max(Threading.Interlocked.Increment(eventId), eventId - 1))

                Collectors.Add(New Collector() With {.SerialNumber = SerialNumberFromDB,
                                                     .ComPortName = ComPortNameFromDB,
                                                     .BaudRate = BaudRateFromDB,
                                                     .Type = CollectorTypeFromDB,
                                                     .Address = AddressFromDB,
                                                     .Group = CollectorGroupFromDB,
                                                     .Description = DescriptionFromDB})
            End While
            ' закрываем ридер
            reader.Close()
        Catch exdb As Exception
            EventLog1.WriteEntry("Ошибка чтения БД: " + exdb.Message, EventLogEntryType.Error, Math.Max(Threading.Interlocked.Increment(eventId), eventId - 1))
            dbConn.Close()
            GetDataIsProgress = False
            Exit Sub
        End Try
        dbConn.Close()
        ' Начинаем опрашивать счетчики
        For Each Collector As Collector In Collectors
            Try
                ComPortName = Collector.ComPortName
                Address = Collector.Address
                CollectorType = Collector.Type
                BaudRate = Collector.BaudRate
                EventLog1.WriteEntry("Компорт: " & ComPortName & "   Адрес счетчика: " & Address & " Скорость порта: " & BaudRate & " Тип счетчика: " & CollectorType, EventLogEntryType.Information, Math.Max(Threading.Interlocked.Increment(eventId), eventId - 1))
                If CollectorType = 1 Then

                    If Not SerialPort.IsOpen Then
                        ' Открываем компорт
                        SerialPort.PortName = ComPortName
                        SerialPort.ReadTimeout = 5000
                        SerialPort.BaudRate = BaudRate
                        Try
                            SerialPort.Open()
                            SerialPort.DiscardInBuffer()
                            SerialPort.DiscardOutBuffer()
                        Catch ex As Exception
                            EventLog1.WriteEntry("Ошибка открытия порта " & ComPortName, EventLogEntryType.Error, Math.Max(Threading.Interlocked.Increment(eventId), eventId - 1))
                            SerialPort.Close()
                            Continue For
                        End Try
                    End If

                    AnswerisOk = MercuryAuth(Address, False)
                    If AnswerisOk = True Then
                        SerialNumber = GetMercurySerialNumber(Address)
                        Dim Ratio As Integer() = GetMercuryRatio(Address)
                        Dim SummaryEnergy As Integer() = GetMercuryEnergy(Address, AskSummaryEnergy)
                        Dim YearEnergy As Integer() = GetMercuryEnergy(Address, AskYearEnergy)
                        Dim MonthEnergy As Integer() = GetMercuryEnergy(Address, AskMonthEnergy)
                        Dim DayEnergy As Integer() = GetMercuryEnergy(Address, AskDayEnergy)
                        'Синхронизация времени раз в сутки в 0 часов и с 0 до 9 минут
                        If Now.Hour = 0 And Now.Minute >= 0 And Now.Minute < 10 Then
                            DoSyncTimeMercury(Address, False)
                        End If

                        EventLog1.WriteEntry(Now.ToString _
                                     & " Компорт: " & ComPortName & vbCrLf _
                                     & " Адрес счетчика: " & Address & vbCrLf _
                                     & " Серийный номер: " & SerialNumber & vbCrLf _
                                     & " Активная энергия: " & SummaryEnergy(0) & vbCrLf _
                                     & " Реактивная энергия: " & SummaryEnergy(1) & vbCrLf _
                                     & " Активная энергия на начало года: " & YearEnergy(0) & vbCrLf _
                                     & " Реактивная энергия на начало года: " & YearEnergy(1) & vbCrLf _
                                     & " Активная энергия на начало месяца: " & MonthEnergy(0) & vbCrLf _
                                     & " Реактивная энергия на начало месяца: " & MonthEnergy(1) & vbCrLf _
                                     & " Активная энергия на начало суток: " & DayEnergy(0) & vbCrLf _
                                     & " Реактивная энергия на начало суток: " & DayEnergy(1) & vbCrLf _
                                     & " Коэффициент трансформации напряжения: " & Ratio(0) & vbCrLf _
                                     & " Коэффициент трансформации тока: " & Ratio(1) & vbCrLf _
                                     , EventLogEntryType.Information, Math.Max(Threading.Interlocked.Increment(eventId), eventId - 1))

                        SavetoBase(ComPortName, Address, SerialNumber, (SummaryEnergy(0) * Ratio(0) * Ratio(1)) / 1000, (SummaryEnergy(1) * Ratio(0) * Ratio(1)) / 1000, (YearEnergy(0) * Ratio(0) * Ratio(1)) / 1000, (YearEnergy(1) * Ratio(0) * Ratio(1)) / 1000,
                                   (MonthEnergy(0) * Ratio(0) * Ratio(1)) / 1000, (MonthEnergy(1) * Ratio(0) * Ratio(1)) / 1000, (DayEnergy(0) * Ratio(0) * Ratio(1)) / 1000, (DayEnergy(1) * Ratio(0) * Ratio(1)) / 1000, Ratio(0), Ratio(1), 1)
                        SerialPort.Close()
                    End If
                Else
                    If Not ModbusPort.IsOpen Then
                        ' Открываем компорт
                        ModbusPort.PortName = ComPortName
                        ModbusPort.ReadTimeout = 5000
                        ModbusPort.BaudRate = BaudRate
                        Try
                            ModbusPort.Open()
                            ModbusPort.DiscardInBuffer()
                            ModbusPort.DiscardOutBuffer()
                        Catch ex As Exception
                            EventLog1.WriteEntry("Ошибка открытия порта " & ComPortName, EventLogEntryType.Error, Math.Max(Threading.Interlocked.Increment(eventId), eventId - 1))
                            ModbusPort.Close()
                            Continue For
                        End Try
                    End If
                    Try
                        Dim master As Modbus.Device.IModbusSerialMaster = Modbus.Device.ModbusSerialMaster.CreateRtu(ModbusPort)
                        master.Transport.Retries = 0
                        master.Transport.ReadTimeout = 500  'millionsecs

                        If Now.Hour = 0 And Now.Minute >= 0 And Now.Minute < 10 Then
                            Dim DateTimeToSatec As UShort() = {Now.Second, Now.Minute, Now.Hour, Now.Day, Now.Month, Now.Year - 2000}
                            ModbusDataWriter(master, Collector.Address, 4352, DateTimeToSatec)
                        End If
                        Dim SatecActiveEnergy As UShort() = ModbusDataReader(master, Collector.Address, 289, 2)
                        If SatecActiveEnergy.Length <> 0 Then
                            Dim iActiveEnergy = SatecActiveEnergy(1) * 10000 + SatecActiveEnergy(0)
                            Dim SatecReactiveEnergy As UShort() = ModbusDataReader(master, Collector.Address, 293, 2)
                            Dim iReactiveEnergy = SatecReactiveEnergy(1) * 10000 + SatecReactiveEnergy(0)
                            Dim SatecSerial = ModbusDataReader(master, Collector.Address, 46080, 2)
                            SerialNumber = SatecSerial(1) * 10000 + SatecSerial(0)

                            EventLog1.WriteEntry(Now.ToString _
                                                               & " Компорт: " & ComPortName & vbCrLf _
                                                               & " Адрес счетчика: " & Address & vbCrLf _
                                                               & " Серийный номер: " & SerialNumber & vbCrLf _
                                                               & " Активная энергия: " & iActiveEnergy & vbCrLf _
                                                               & " Реактивная энергия: " & iReactiveEnergy & vbCrLf _
                                                                 , EventLogEntryType.Information, Math.Max(Threading.Interlocked.Increment(eventId), eventId - 1))

                            SavetoBase(ComPortName, Address, SerialNumber, iActiveEnergy, iReactiveEnergy, 0, 0, 0, 0, 0, 0, 1, 1, 2)

                            ModbusPort.Close()
                        End If
                    Catch Ex As Exception
                        EventLog1.WriteEntry(Ex.Message, EventLogEntryType.Warning, Math.Max(Threading.Interlocked.Increment(eventId), eventId - 1))
                        ModbusPort.Close()
                        Continue For
                    End Try
                End If
            Catch Ex As Exception
                EventLog1.WriteEntry(Ex.Message, EventLogEntryType.Warning, Math.Max(Threading.Interlocked.Increment(eventId), eventId - 1))
                ModbusPort.Close()
                Continue For
            End Try
        Next
        GetDataIsProgress = False
    End Sub

    Public Function MercuryAuth(ByVal CollectorAddress As String, ByVal Admin As Boolean) As Boolean
        ' Делаем логин в счетчик
        Dim CRC, SendtoCom, ReceiveFromCom, Login As Byte()
        Dim DataLenght As Integer
        Dim Address As Byte

        Address = Byte.Parse(CollectorAddress)
        If Admin Then
            'Как админ
            Login = {Address, 1, 2, 2, 2, 2, 2, 2, 2}
        Else
            ' Как юзер
            Login = {Address, 1, 1, 1, 1, 1, 1, 1, 1}
        End If
        ' Считаем контрольную сумму запроса
        CRC = CRC16(Login)
        ' Объединяем запрос и контрольную сумму
        SendtoCom = Login.Concat(CRC).ToArray
        DataLenght = SendtoCom.Count
        Try
            ' Отправляем запрос
            SendSerialData(SendtoCom, 0, DataLenght)
            ' Читаем данные
            ReceiveFromCom = ReceiveSerialData(4)
        Catch ex As Exception
            EventLog1.WriteEntry("Неправильный логин или пароль " & " Адрес: " & Address, EventLogEntryType.Warning, Math.Max(Threading.Interlocked.Increment(eventId), eventId - 1))
            Return False
        End Try
        Return True

    End Function

    Public Function GetMercurySerialNumber(ByVal CollectorAddress As String) As Integer
        ' Запрос серийного номера счетчика
        Dim AskSerialNumber, ReceiveFromComSerialNumber As Byte()
        Dim SerialNumber As Integer
        Dim Address As Byte
        Address = Byte.Parse(CollectorAddress)
        AskSerialNumber = {Address, 8, 0}
        ReceiveFromComSerialNumber = SendCommandtoMercury(AskSerialNumber, 10)
        If ReceiveFromComSerialNumber.Length <> 0 Then
            SerialNumber = ReceiveFromComSerialNumber(1) * 1000000 + ReceiveFromComSerialNumber(2) * 10000 + ReceiveFromComSerialNumber(3) * 100 + ReceiveFromComSerialNumber(4)
        Else
            EventLog1.WriteEntry("Неправильный порт или адрес " & " Адрес: " & Address, EventLogEntryType.Warning, Math.Max(Threading.Interlocked.Increment(eventId), eventId - 1))
            SerialNumber = 0
        End If
        If SerialNumber = 0 Then
            EventLog1.WriteEntry("Серийный номер равен нулю у счетчика " & Address, EventLogEntryType.Warning, Math.Max(Threading.Interlocked.Increment(eventId), eventId - 1))
        End If
        Return SerialNumber
    End Function

    Public Function GetMercuryRatio(ByVal CollectorAddress As String) As Integer()
        ' Запрос коэффициентов трансформации
        Dim AskRatio, ReceiveFromComRatio As Byte()
        Dim VoltageRatioRaw As Byte() = New Byte(1) {}
        Dim CurrentRatioRaw As Byte() = New Byte(1) {}
        Dim VoltageRatio As Integer
        Dim CurrentRatio As Integer
        Dim Address As Byte
        Address = Byte.Parse(CollectorAddress)
        AskRatio = {Address, 8, 2}
        Dim AskLoop = 0

        Do
            ReceiveFromComRatio = SendCommandtoMercury(AskRatio, 7)
            AskLoop += 1

        Loop Until CheckMercuryAnswer(ReceiveFromComRatio, Byte.Parse(CollectorAddress)) Or AskLoop = 3


        If ReceiveFromComRatio.Length <> 0 Or CheckMercuryAnswer(ReceiveFromComRatio, Address) Then
            VoltageRatioRaw(0) = ReceiveFromComRatio(1)
            VoltageRatioRaw(1) = ReceiveFromComRatio(2)
            CurrentRatioRaw(0) = ReceiveFromComRatio(3)
            CurrentRatioRaw(1) = ReceiveFromComRatio(4)
            'Конвертируем из набора байтов в 16-ричное число
            VoltageRatio = BitConverter.ToInt16(VoltageRatioRaw.Reverse.ToArray, 0)
            CurrentRatio = BitConverter.ToInt16(CurrentRatioRaw.Reverse.ToArray, 0)
        Else
            EventLog1.WriteEntry("Неправильный порт или адрес: " & " Адрес: " & Address, EventLogEntryType.Warning, Math.Max(Threading.Interlocked.Increment(eventId), eventId - 1))
            VoltageRatio = 0
            CurrentRatio = 0
        End If
        Return {VoltageRatio, CurrentRatio}

    End Function

    Public Function GetMercuryEnergy(ByVal CollectorAddress As String, ByVal bAskEnergy As Byte()) As Integer()
        Dim Ask, ReceiveFromComEnergy As Byte()
        Dim ActiveEnergyRaw As Byte() = New Byte(3) {}
        Dim ReactiveEnergyRaw As Byte() = New Byte(3) {}
        Dim ActiveEnergy, ReactiveEnergy As UInteger
        Dim Address As Byte() = New Byte(0) {}

        Address(0) = Byte.Parse(CollectorAddress)
        Ask = Address.Concat(bAskEnergy).ToArray
        Dim AskLoop = 0
        Do
            ReceiveFromComEnergy = SendCommandtoMercury(Ask, 19)
            AskLoop += 1

        Loop Until CheckMercuryAnswer(ReceiveFromComEnergy, Byte.Parse(CollectorAddress)) Or AskLoop = 3
        If CheckMercuryAnswer(ReceiveFromComEnergy, Byte.Parse(CollectorAddress)) Then

            ActiveEnergyRaw(0) = ReceiveFromComEnergy(2)
            ActiveEnergyRaw(1) = ReceiveFromComEnergy(1)
            ActiveEnergyRaw(2) = ReceiveFromComEnergy(4)
            ActiveEnergyRaw(3) = ReceiveFromComEnergy(3)

            ReactiveEnergyRaw(0) = ReceiveFromComEnergy(10)
            ReactiveEnergyRaw(1) = ReceiveFromComEnergy(9)
            ReactiveEnergyRaw(2) = ReceiveFromComEnergy(12)
            ReactiveEnergyRaw(3) = ReceiveFromComEnergy(11)
            ActiveEnergy = BitConverter.ToUInt32(ActiveEnergyRaw.Reverse.ToArray, 0)
            ReactiveEnergy = BitConverter.ToUInt32(ReactiveEnergyRaw.Reverse.ToArray, 0)
        Else
            ActiveEnergy = 0
            ReactiveEnergy = 0
        End If

        Return {ActiveEnergy, ReactiveEnergy}

    End Function

    Public Function GetDateTimeFromMercury(ByVal bAddress As Byte) As DateTime
        'Запрос даты и времени из счетчика
        Dim AskDatetime, ReceiveFromComDateTime As Byte()
        Dim DateTimeFromCollector As DateTime = "01-01-2000"
        Dim bCurrentDateTimeFromCollectorRaw As Byte() = New Byte(8) {}

        AskDatetime = {bAddress, 4, 0}

        ReceiveFromComDateTime = SendCommandtoMercury(AskDatetime, 11)
        If ReceiveFromComDateTime.Length <> 0 Or CheckMercuryAnswer(ReceiveFromComDateTime, bAddress) Then
            'секунды, минуты, часы, день недели, число, месяц, год, признак зима/лето
            Array.Copy(ReceiveFromComDateTime, 1, bCurrentDateTimeFromCollectorRaw, 0, 8)

            DateTimeFromCollector = DateTimeFromCollector.AddYears(CInt(Hex(bCurrentDateTimeFromCollectorRaw(6)))) _
                                                         .AddMonths(CInt(Hex(bCurrentDateTimeFromCollectorRaw(5))) - 1) _
                                                         .AddDays(CInt(Hex(bCurrentDateTimeFromCollectorRaw(4))) - 1) _
                                                         .AddHours(CInt(Hex(bCurrentDateTimeFromCollectorRaw(2)))) _
                                                         .AddMinutes(CInt(Hex(bCurrentDateTimeFromCollectorRaw(1)))) _
                                                         .AddSeconds(CInt(Hex(bCurrentDateTimeFromCollectorRaw(0))))

        Else
            EventLog1.WriteEntry("Не удалось прочитать время " & " Адрес: " & bAddress, EventLogEntryType.Warning, Math.Max(Threading.Interlocked.Increment(eventId), eventId - 1))

        End If
        Return DateTimeFromCollector
    End Function

    Public Function SendCommandtoMercury(Message As Byte(), ReceiveDataLenght As Integer) As Byte()

        Dim CRC, SendtoCom, ReceiveFromCom As Byte()
        Dim DataLenght As Integer
        Dim Address As Byte = Message(0)

        CRC = CRC16(Message)
        SendtoCom = Message.Concat(CRC).ToArray()
        DataLenght = SendtoCom.Count
        SendSerialData(SendtoCom, 0, DataLenght)
        ReceiveFromCom = ReceiveSerialData(ReceiveDataLenght)
        Return ReceiveFromCom

    End Function

    Public Function CRC16(data As Byte()) As Byte()
        Dim crcfull As UInt16 = &HFFFF
        Dim crchigh As Byte, crclow As Byte
        Dim crclsb As Byte

        For i As Integer = 0 To data.Length - 1
            crcfull = crcfull Xor data(i)

            For j As Integer = 0 To 7
                crclsb = crcfull And &H1
                crcfull >>= 1

                If (crclsb <> 0) Then
                    crcfull = crcfull Xor &HA001
                End If
            Next
        Next

        crchigh = (crcfull >> 8) And &HFF
        crclow = crcfull And &HFF
        Return New Byte(1) {crclow, crchigh}
    End Function

    Public Sub SendSerialData(ByVal data As Byte(), ByVal offset As Integer, ByVal count As Integer)

        Try
            SerialPort.Write(data, offset, count)
        Catch ex As Exception
            EventLog1.WriteEntry("Ошибка отправки данных в порт ", EventLogEntryType.Error, Math.Max(Threading.Interlocked.Increment(eventId), eventId - 1))
            Return
        End Try

    End Sub

    Public Function ReceiveSerialData(ByVal DataLenght As Integer) As Byte()

        ' Receive strings from a serial port.
        Dim ReturnByte As Byte() = New Byte(DataLenght - 1) {}

        Try
            SerialPort.Read(ReturnByte, 0, DataLenght)
        Catch ex As Exception
            EventLog1.WriteEntry(ex.Message.ToString, EventLogEntryType.Warning, Math.Max(Threading.Interlocked.Increment(eventId), eventId - 1))
            Return {}
        End Try
        Return ReturnByte
    End Function

    Public Sub DoSyncTimeMercury(ByVal bAddress As Byte, ByVal Silent As Boolean)
        Dim DateTimeCollector As DateTime
        Dim Seconds, Minutes, Hours As Byte
        Dim SetDatetime, ReceiveFromCom As Byte()
        DateTimeCollector = GetDateTimeFromMercury(bAddress)
        If Date.Compare(Now.AddMinutes(-4), DateTimeCollector) > 0 Or Date.Compare(Now.AddMinutes(+4), DateTimeCollector) < 0 Then
            EventLog1.WriteEntry("Корректировка счетчика " & SerialPort.PortName.ToString & " " & bAddress.ToString & " более 4 мин. невозможна. Установите время с помощью конфигуратора", EventLogEntryType.Information, Math.Max(Threading.Interlocked.Increment(eventId), eventId - 1))
        Else
            Seconds = Convert.ToInt32(DateTimeCollector.Second.ToString, 16)
            Minutes = Convert.ToInt32(DateTimeCollector.Minute.ToString, 16)
            Hours = Convert.ToInt32(DateTimeCollector.Hour.ToString, 16)
            SetDatetime = {bAddress, 3, &HD, Seconds, Minutes, Hours}
            ReceiveFromCom = SendCommandtoMercury(SetDatetime, 4)
            If Not Silent Then
                Select Case ReceiveFromCom(1)
                    Case 0
                        EventLog1.WriteEntry("Счетчик " & SerialPort.PortName.ToString & " " & bAddress.ToString & " синхронизирован", EventLogEntryType.Information, Math.Max(Threading.Interlocked.Increment(eventId), eventId - 1))
                    Case 4
                        EventLog1.WriteEntry("Счетчик " & SerialPort.PortName.ToString & " " & bAddress.ToString & " уже был синхронизирован в течение этих суток", EventLogEntryType.Information, Math.Max(Threading.Interlocked.Increment(eventId), eventId - 1))
                        Exit Select
                    Case Is <> 0
                        EventLog1.WriteEntry("Ошибка синхронизации времени в счетчике " & SerialPort.PortName.ToString & " " & bAddress.ToString, EventLogEntryType.Information, Math.Max(Threading.Interlocked.Increment(eventId), eventId - 1))
                End Select
            End If
        End If
    End Sub

    Public Function ModbusDataReader(master As Modbus.Device.IModbusSerialMaster, SlaveAddress As String, StartRegister As String, PointsNumber As UShort) As UShort()

        Dim holding_register As UShort()
        Dim slaveId As Byte = Byte.Parse(SlaveAddress)
        Dim startAddress As UShort = UShort.Parse(StartRegister)

        Try
            holding_register = master.ReadHoldingRegisters(slaveId, startAddress, PointsNumber)
        Catch ex As Exception
            holding_register = {}
            MsgBox(ex.Message & " Неправильный порт или адрес " & " Адрес: " & SlaveAddress)
        End Try

        Return holding_register
    End Function

    Public Sub ModbusDataWriter(master As Modbus.Device.IModbusSerialMaster, SlaveAddress As String, StartRegister As String, WriteData As UShort())
        Dim slaveId As Byte = Byte.Parse(SlaveAddress)
        Dim startAddress As UShort = UShort.Parse(StartRegister)
        Try
            master.WriteMultipleRegisters(slaveId, startAddress, WriteData)
        Catch ex As Exception
            MsgBox(ex.Message & " Неправильный порт или адрес " & " Адрес: " & SlaveAddress)
        End Try
    End Sub

    Public Function CheckMercuryAnswer(ByVal Answer As Byte(), CollectorNumber As Byte) As Boolean
        Dim CRC As Byte() = {0, 0}
        Dim Data = Answer.Clone()
        Array.Copy(Data, Data.Length - 2, CRC, 0, 2)
        Array.Resize(Data, Data.Length - 2)
        Dim CalcCRC As Byte() = CRC16(Data)
        Return (CRC.SequenceEqual(CalcCRC) And Answer(0) = CollectorNumber)
    End Function

    Public Function SavetoBase(Port As String, Address As String, SerialNumber As Integer, ActiveEnergy As UInteger, ReactiveEnergy As UInteger, YearActiveEnergy As UInteger, YearReactiveEnergy As UInteger,
                   MonthActiveEnergy As UInteger, MonthReactiveEnergy As UInteger, DayActiveEnergy As UInteger, DayReactiveEnergy As UInteger, VoltageRatio As Integer, CurrentRatio As Integer, CollectorType As Integer) As Boolean
        Dim dbConn As New OleDbConnection()
        ' Подключаемся к базе для записи данных
        If ((ActiveEnergy = 0 Or ReactiveEnergy = 0 Or YearActiveEnergy = 0 Or YearReactiveEnergy = 0 Or MonthActiveEnergy = 0 Or MonthReactiveEnergy = 0 Or DayActiveEnergy = 0 _
            Or DayReactiveEnergy = 0 Or VoltageRatio = 0 Or CurrentRatio = 0) And CollectorType = 1) Or ((ActiveEnergy = 0 Or ReactiveEnergy = 0) And CollectorType = 2) Then
            EventLog1.WriteEntry("Ошибка получения данных " & " Активная энергия: " & ActiveEnergy & " Реактивная энергия: " & ReactiveEnergy & " Активная энергия за год: " _
                                 & YearActiveEnergy & " Реактивная энергия за год: " & YearReactiveEnergy & " Активная энергия за месяц: " & MonthActiveEnergy & " Реактивная энергия за месяц: " _
                                 & MonthReactiveEnergy & " Активная энергия за день: " & DayActiveEnergy & " Реактивная энергия за день: " & DayReactiveEnergy _
                                 & " Коэффициент трансформации напряжения: " & VoltageRatio & " Коэффициент трансформации тока: " & CurrentRatio _
                                 , EventLogEntryType.Warning, Math.Max(Threading.Interlocked.Increment(eventId), eventId - 1))
            Return False
        End If
        Try
            dbConn.ConnectionString = "Provider=SQLOLEDB; Data Source=192.168.10.32\SQLEXPRESS; Initial Catalog=Energy; Persist Security Info=False; User Id=energy; Password=energy"
            dbConn.Open()
        Catch ex As Exception
            EventLog1.WriteEntry("Ошибка соединения с базой данных: " + ex.Message, EventLogEntryType.Error, Math.Max(Threading.Interlocked.Increment(eventId), eventId - 1))
            Return False
        End Try

        Try
            ' Делаем запрос
            Dim dbCommand As OleDbCommand = dbConn.CreateCommand()
            dbCommand.CommandText = "insert into AccuEnergy ([datetime], [Port],[Address],[SerialNumber],[SummaryActiveEnergy],[SummaryReactiveEnergy],[YearActiveEnergy], [YearReactiveEnergy], [MonthActiveEnergy], [MonthReactiveEnergy], [DayActiveEnergy], [DayReactiveEnergy], [VoltageRatio], [CurrentRatio], [CollectorType]) VALUES (getdate(), ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
            dbCommand.Parameters.Add("Port", OleDbType.[VarChar]).Value = Port
            dbCommand.Parameters.Add("Address", OleDbType.[VarChar]).Value = Address
            dbCommand.Parameters.Add("SerialNumber", OleDbType.[Integer]).Value = SerialNumber
            dbCommand.Parameters.Add("SummaryActiveEnergy", OleDbType.UnsignedInt).Value = ActiveEnergy
            dbCommand.Parameters.Add("SummaryReactiveEnergy", OleDbType.UnsignedInt).Value = ReactiveEnergy
            dbCommand.Parameters.Add("YearActiveEnergy", OleDbType.UnsignedInt).Value = YearActiveEnergy
            dbCommand.Parameters.Add("YearReactiveEnergy", OleDbType.UnsignedInt).Value = YearReactiveEnergy
            dbCommand.Parameters.Add("MonthActiveEnergy", OleDbType.UnsignedInt).Value = MonthActiveEnergy
            dbCommand.Parameters.Add("MonthReactiveEnergy", OleDbType.UnsignedInt).Value = MonthReactiveEnergy
            dbCommand.Parameters.Add("DayActiveEnergy", OleDbType.UnsignedInt).Value = DayActiveEnergy
            dbCommand.Parameters.Add("DayReactiveEnergy", OleDbType.UnsignedInt).Value = DayReactiveEnergy
            dbCommand.Parameters.Add("VoltageRatio", OleDbType.Integer).Value = VoltageRatio
            dbCommand.Parameters.Add("CurrentRatio", OleDbType.Integer).Value = CurrentRatio
            dbCommand.Parameters.Add("CollectorType", OleDbType.Integer).Value = CollectorType
            dbCommand.ExecuteNonQuery()
            dbConn.Close()
            Return True
        Catch exdb As Exception
            EventLog1.WriteEntry("Ошибка записи в БД: " + exdb.Message, EventLogEntryType.Error, Math.Max(Threading.Interlocked.Increment(eventId), eventId - 1))
            dbConn.Close()
            Return False
        End Try
        ' Закрываем соединение с базой
        dbConn.Close()
        Return False
    End Function
End Class

Friend Class Collector
    ' Класс для сохранения данных счетчика
    Public Property SerialNumber As Integer
    Public Property ComPortName As String
    Public Property BaudRate As Integer
    Public Property Address As String
    Public Property Group As Integer
    Public Property Description As String
    Public Property Type As String
End Class


