Imports System.Threading
Imports System.Text
Imports System.Net.Sockets
Imports System.Net
Imports System.Text.RegularExpressions
Imports HomeSeerAPI.VSVGPairs
Imports HomeSeerAPI
Imports System.IO
Imports System.ComponentModel

Public Class hspi_audio_device

  Dim WithEvents serialPort As New IO.Ports.SerialPort

  Protected RefreshStateThread As Thread
  Protected SendCommandThread As Thread
  Protected WatchdogThread As Thread

  Protected gAudioZones As New List(Of hspi_audio_zone)

  Protected gCommandQueue As New Queue
  Protected gDelayCounter As UInteger = 0

  Protected m_DeviceId As Integer = 0
  Protected m_DeviceFeatures As hspi_audio_device_features

  Protected m_ConnectionType As String = ""
  Protected m_ConnectionAddr As String = ""
  Protected m_DeviceLastResponse As DateTime = DateTime.Now

  Protected m_Initialized As Boolean = False
  Protected m_WatchdogActive As Boolean = False
  Protected m_WatchdogDisabled As Boolean = False

  Protected m_DeviceName As String = ""
  Protected m_DeviceSerial As String = ""
  Protected m_DeviceMake As String = ""
  Protected m_DeviceModel As String = ""
  Protected m_DeviceRevision As String = "1.0"
  Protected m_DeviceZones As Integer = 0
  Protected m_DeviceTunerSource As Integer = 0

  Protected gDeviceInitialized As Boolean = False
  Protected gDeviceConnected As Boolean = False
  Protected gDevicePoweredOn As Boolean = False
  Protected gDeviceResponse As Boolean = False
  Protected gDeviceRefreshSystemState As Boolean = False

#Region "Audio Device Object"

  Public ReadOnly Property DeviceId() As Integer
    Get
      Return Me.m_DeviceId
    End Get
  End Property

  Public Property DeviceName() As String
    Get
      Return Me.m_DeviceName
    End Get
    Set(value As String)
      Me.m_DeviceName = value
    End Set
  End Property

  Public Property DeviceSerial() As String
    Get
      Return Me.m_DeviceSerial
    End Get
    Set(value As String)
      Me.m_DeviceSerial = value
    End Set
  End Property

  Public Property ConnectionType() As String
    Get
      Return Me.m_ConnectionType
    End Get
    Set(value As String)
      Me.m_ConnectionType = value
    End Set
  End Property

  Public Property ConnectionAddr() As String
    Get
      Return Me.m_ConnectionAddr
    End Get
    Set(value As String)
      Me.m_ConnectionAddr = value
    End Set
  End Property

  Public Property DeviceMake() As String
    Get
      Return Me.m_DeviceMake
    End Get
    Set(value As String)
      Me.m_DeviceMake = value
    End Set
  End Property

  Public Property DeviceModel() As String
    Get
      Return Me.m_DeviceModel
    End Get
    Set(value As String)
      Me.m_DeviceModel = value
    End Set
  End Property

  Public Property DeviceRevision() As String
    Get
      Return Me.m_DeviceRevision
    End Get
    Set(value As String)
      Me.m_DeviceRevision = value
    End Set
  End Property

  Public Property DeviceZones() As Integer
    Get
      Return Me.m_DeviceZones
    End Get
    Set(value As Integer)
      Me.m_DeviceZones = value
    End Set
  End Property

  Public Property DeviceTunerSource() As Integer
    Get
      Return Me.m_DeviceTunerSource
    End Get
    Set(value As Integer)
      Me.m_DeviceTunerSource = value
    End Set
  End Property

  Public ReadOnly Property DevicePowerStatus() As String
    Get
      Select Case gDevicePoweredOn
        Case True
          Return "On"
        Case Else
          Return "Off"
      End Select
    End Get
  End Property

  Public ReadOnly Property ConnectionStatus() As String
    Get
      Select Case gDeviceConnected
        Case True
          Return "Connected"
        Case Else
          Return "Disconnected"
      End Select
    End Get
  End Property

  Public ReadOnly Property AudioZones() As List(Of hspi_audio_zone)
    Get
      Return gAudioZones
    End Get
  End Property

  Public Property DeviceLastResponse() As DateTime
    Set(ByVal value As DateTime)
      m_DeviceLastResponse = value
    End Set
    Get
      Return m_DeviceLastResponse
    End Get
  End Property

  Public Sub New(ByVal DeviceId As Integer,
                 ByVal strDeviceName As String,
                 ByVal strDeviceSerial As String,
                 ByVal strConnectionType As String,
                 ByVal strConnectionAddr As String,
                 ByVal strMake As String,
                 ByVal strModel As String,
                 ByVal iDeviceZones As Integer,
                 ByVal iDeviceTunerSource As Integer)

    MyBase.New()

    Try
      '
      ' Set the device_id for this object
      '
      Me.m_DeviceId = DeviceId

      Me.m_DeviceName = strDeviceName
      Me.m_DeviceSerial = strDeviceSerial

      Me.m_ConnectionType = strConnectionType
      Me.m_ConnectionAddr = strConnectionAddr

      Me.m_DeviceMake = strMake
      Me.m_DeviceModel = strModel
      Me.m_DeviceZones = iDeviceZones
      Me.m_DeviceTunerSource = iDeviceTunerSource

      m_DeviceFeatures = New hspi_audio_device_features(Me.DeviceModel)

      Dim strMessage As String = ""

      Dim KeyTypes() As String = GetAudioZoneKeyTypes()

      '
      ' Initialize the Audio Zone Devices
      '
      For Each strKeyType As String In KeyTypes
        Select Case strKeyType
          Case "ST2 Smart Tuner"
            '
            ' Process a ST2 Smart Tuner
            '
            Dim AudioZoneKeyNames() As String = GetAudioZoneKeyNames(strKeyType)

            For Each AudioZoneKeyName As String In AudioZoneKeyNames
              For cc As Integer = 0 To 0
                For tt As Integer = 0 To 1
                  Dim audioZoneId As String = String.Format("{0}.{1}.{2}", m_DeviceId, cc, tt)
                  Dim dv_addr As String = String.Format("Russound{0}-{1}", audioZoneId, AudioZoneKeyName)
                  If gAudioZones.Exists(Function(s) s.DeviceAddr = dv_addr) = False Then
                    gAudioZones.Add(New hspi_audio_zone(m_DeviceId, cc, tt, dv_addr))
                  End If
                Next
              Next

            Next

          Case Else
            '
            ' Process an Audio Zone
            '
            Dim AudioZoneKeyNames() As String = GetAudioZoneKeyNames(strKeyType)

            For Each AudioZoneKeyName As String In AudioZoneKeyNames
              For cc As Integer = 0 To 5
                For zz As Integer = 0 To 5
                  Dim audioZoneId As String = String.Format("{0}.{1}.{2}", m_DeviceId, cc, zz)
                  Dim dv_addr As String = String.Format("Russound{0}-{1}", audioZoneId, AudioZoneKeyName)
                  If gAudioZones.Exists(Function(s) s.DeviceAddr = dv_addr) = False Then
                    gAudioZones.Add(New hspi_audio_zone(m_DeviceId, cc, zz, dv_addr))
                  End If
                Next
              Next

            Next

        End Select

      Next

      '
      ' Start the process command queue thread
      '
      SendCommandThread = New Thread(New ThreadStart(AddressOf ProcessCommandQueue))
      SendCommandThread.Name = "ProcessCommandQueue"
      SendCommandThread.Start()

      strMessage = SendCommandThread.Name & " Thread Started"
      WriteMessage(strMessage, MessageType.Debug)

      '
      ' Start the process command queue thread
      '
      RefreshStateThread = New Thread(New ThreadStart(AddressOf RefreshSystemState))
      RefreshStateThread.Name = "RefreshState"
      RefreshStateThread.Start()

      strMessage = RefreshStateThread.Name & " Thread Started"
      WriteMessage(strMessage, MessageType.Debug)

      '
      ' Start the watchdog thread
      '
      WatchdogThread = New Thread(New ThreadStart(AddressOf ConnectionWatchdogThread))
      WatchdogThread.Name = "WatchdogThread"
      WatchdogThread.Start()

      strMessage = WatchdogThread.Name & " Thread Started"
      WriteMessage(strMessage, MessageType.Debug)

    Catch pEx As Exception
      WriteMessage(pEx.Message, MessageType.Error)
    End Try

  End Sub

  Protected Overrides Sub Finalize()

    Try

      '
      ' Abort SendCommandThread
      '
      If SendCommandThread.IsAlive = True Then
        SendCommandThread.Abort()
      End If

      '
      ' Abort WatchdogThread
      '
      If WatchdogThread.IsAlive = True Then
        WatchdogThread.Abort()
      End If

      '
      ' Abort RefreshStateThread
      '
      If RefreshStateThread.IsAlive = True Then
        RefreshStateThread.Abort()
      End If

      DisconnectFromDevice()

    Catch pEx As Exception

    End Try

    MyBase.Finalize()

  End Sub

#End Region

#Region "Audio Device - Watchdog"

  ''' <summary>
  ''' Thread that checks we are still connected to the Audio Device
  ''' </summary>
  ''' <remarks></remarks>
  Public Sub ConnectionWatchdogThread()

    Dim bAbortThread As Boolean = False
    Dim strMessage As String = ""
    Dim dblTimerInterval As Single = 1000 * 30
    Dim iSeconds As Long = 0

    Try
      '
      ' Stay in Connection Watchdog Thread for duration of program
      '
      While bAbortThread = False

        Try

          If m_WatchdogDisabled = True Then

            dblTimerInterval = 1000 * 60

            strMessage = String.Format("Watchdog Timer indicates the Audio Device device '{0}' auto reconnect is disabled.", m_ConnectionAddr)
            WriteMessage(strMessage, MessageType.Debug)

          ElseIf gDeviceInitialized = True Then

            If IsDate(m_DeviceLastResponse) Then
              iSeconds = DateDiff(DateInterval.Second, m_DeviceLastResponse, DateTime.Now)
            End If

            strMessage = String.Format("Watchdog Timer indicates a response from the Audio Device device '{0}' was received at {1}.", m_ConnectionAddr, m_DeviceLastResponse.ToString)
            WriteMessage(strMessage, MessageType.Debug)

            '
            ' Test to see if we are connected and that we have received a response within the past 300 seconds
            '
            Call CheckPhysicalConnection()

            If iSeconds > 300 Or m_WatchdogActive = True Or gDeviceConnected = False Then
              '
              ' Action for initial watchdog trigger
              '
              If m_WatchdogActive = False Then
                m_WatchdogActive = True
                dblTimerInterval = 1000 * 30

                Dim strWatchdogReason As String = String.Format("No response response from the Audio Device device '{0}' for {1} seconds.", m_ConnectionAddr, iSeconds)
                If gDeviceConnected = False Then
                  strWatchdogReason = String.Format("Connection to Audio Device device '{0}' was lost.", m_ConnectionAddr)
                End If

                strMessage = String.Format("Watchdog Timer indicates {0}.  Attempting to reconnect ...", strWatchdogReason)
                WriteMessage(strMessage, MessageType.Warning)

                '
                ' Check watchdog trigger
                '
                Dim strTrigger As String = IFACE_NAME & Chr(2) & "Audio Device Watchdog Trigger" & Chr(2) & "Connection Failure" & Chr(2) & "*"
                'callback.CheckTrigger(strTrigger)
              End If

              '
              ' Ensure everything is closed properly and attempt a reconnect
              '
              Call DisconnectFromDevice()
              Call ConnectToDevice()

              If gDeviceConnected = False Then

                WriteMessage("Watchdog Timer reconnect attempt failed.", MessageType.Warning)

                dblTimerInterval *= 2
                If dblTimerInterval > 3600000 Then
                  dblTimerInterval = 3600000
                End If

              Else

                WriteMessage("Watchdog Timer reconnect attempt succeeded.", MessageType.Informational)
                m_WatchdogActive = False
                dblTimerInterval = 1000 * 30

                '
                ' Check watchdog trigger
                '
                Dim strTrigger As String = IFACE_NAME & Chr(2) & "Audio Device Watchdog Trigger" & Chr(2) & "Connection Restore" & Chr(2) & "*"
                'callback.CheckTrigger(strTrigger)

                Call CheckPowerStatus()

              End If

            Else
              '
              ' Plug-in is connected to the Global Cache device
              '
              m_WatchdogActive = False
              dblTimerInterval = 1000 * 30

              strMessage = String.Format("Watchdog Timer indicates a response from the Audio Device device '{0}' was received {1} seconds ago.", m_ConnectionAddr, iSeconds.ToString)
              WriteMessage(strMessage, MessageType.Debug)

              Call CheckPowerStatus()

            End If

          End If

          '
          ' Sleep Watchdog Thread
          '
          strMessage = String.Format("Watchdog Timer thread for the Audio Device device '{0}' sleeping for {1}.", m_ConnectionAddr, dblTimerInterval.ToString)
          WriteMessage(strMessage, MessageType.Debug)

          Thread.Sleep(dblTimerInterval)

        Catch pEx As ThreadInterruptedException
          '
          ' Thread sleep was interrupted
          '
          gDeviceInitialized = True
          strMessage = String.Format("Watchdog Timer thread for the Audio Device device '{0}' was interrupted.", m_ConnectionAddr, iSeconds.ToString)
          WriteMessage(strMessage, MessageType.Debug)

        Catch pEx As Exception
          '
          ' Process Exception
          '
          Call ProcessError(pEx, "ConnectionWatchdogThread()")
        End Try

      End While ' Stay in thread until we get an abort/exit request

    Catch ab As ThreadAbortException
      '
      ' Process Thread Abort Exception
      '
      bAbortThread = True      ' Not actually needed
      Call WriteMessage("Abort requested on ConnectionWatchdogThread", MessageType.Debug)
    Catch pEx As Exception
      '
      ' Process Exception
      '
      Call ProcessError(pEx, "ConnectionWatchdogThread()")
    Finally

    End Try

  End Sub

#End Region

#Region "Audio Device Connection Initilization/Shutdown"

  ''' <summary>
  ''' Initialize the connection to the Audio Device
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function ConnectToDevice() As String

    Dim strMessage As String = ""

    strMessage = "Entered ConnectToDevice() function."
    Call WriteMessage(strMessage, MessageType.Debug)

    Try
      '
      ' Get the Connection Information
      '
      Dim DeviceDevice As Hashtable = GetAudioDevice(Me.m_DeviceId)

      m_ConnectionType = DeviceDevice("device_conn")
      m_ConnectionAddr = DeviceDevice("device_addr")

      Select Case m_ConnectionType.ToUpper
        Case "ETHERNET"
          '
          ' Try connection via the Audio Device
          '
          strMessage = String.Format("Initiating Audio Device '{0}' connection ...", m_ConnectionAddr)
          Call WriteMessage(strMessage, MessageType.Debug)

          '
          ' Inititalize Ethernet connection
          '
          Dim regexPattern As String = "(?<ipaddr>.+):(?<port>\d+)"
          If Regex.IsMatch(m_ConnectionAddr, regexPattern) = True Then

            Dim ip_addr As String = Regex.Match(m_ConnectionAddr, regexPattern).Groups("ipaddr").ToString()
            Dim ip_port As String = Regex.Match(m_ConnectionAddr, regexPattern).Groups("port").ToString()

            gDeviceConnected = ConnectToEthernet(ip_addr, ip_port)
            If gDeviceInitialized = False Then
              gDeviceInitialized = gDeviceConnected
            End If

          Else
            '
            ' Unable to connect
            '
            gDeviceConnected = False

          End If

          If gDeviceConnected = False Then
            Throw New Exception(String.Format("Unable to connect to Audio Device '{0}'.", m_ConnectionAddr))
          End If

        Case "SERIAL"
          '
          ' Try connecting to the serial port
          '
          Dim strComPort As String = m_ConnectionAddr.Replace(":", "").ToUpper

          strMessage = String.Format("Initiating Audio Device '{0}' connection ...", m_ConnectionAddr)
          Call WriteMessage(strMessage, MessageType.Debug)

          '
          ' Close port if already open
          '
          If serialPort.IsOpen Then
            serialPort.Close()
          End If

          Try
            With serialPort
              .PortName = strComPort
              .BaudRate = 19200
              .Parity = IO.Ports.Parity.None
              .DataBits = 8
              .StopBits = IO.Ports.StopBits.One
              .ReadTimeout = 100
            End With
            serialPort.Open()
            serialPort.DiscardInBuffer()

            gDeviceConnected = serialPort.IsOpen
            If gDeviceInitialized = False Then
              gDeviceInitialized = gDeviceConnected
            End If

          Catch pEx As Exception
            Throw New Exception(String.Format("Unable to connect to Audio Device device '{0}' because '{1}'.", m_ConnectionAddr, pEx.ToString))
          End Try

        Case Else
          '
          ' Bail out when no port is defined (user has not set a port yet)
          '
          strMessage = String.Format("Audio Device device '{0}' interface is disabled.", Me.ConnectionAddr)
          Call WriteMessage(strMessage, MessageType.Warning)

          '
          ' Unable to connect because the Interface is disabled
          '
          Throw New Exception(strMessage)

      End Select

      gDeviceRefreshSystemState = True

      '
      ' We are connected here
      '
      Return ""

    Catch pEx As Exception
      '
      ' Process program exception
      '
      gDeviceConnected = False
      Return pEx.ToString
    Finally
      '
      ' Update the Audio Device connection status
      '
      UpdateAudioConnectionDevice(Me.m_ConnectionType, Me.m_DeviceId, gDeviceConnected)
    End Try

  End Function

  ''' <summary>
  ''' Disconnect the connection to Audio Device
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function DisconnectFromDevice() As Boolean

    Dim strMessage As String = ""

    strMessage = "Entered DisconnectFromDevice() function."
    Call WriteMessage(strMessage, MessageType.Debug)

    Try

      Select Case m_ConnectionType.ToUpper
        Case "ETHERNET"
          '
          ' Close the ethernet connection
          '
          DisconnectEthernet()
        Case "SERIAL"
          '
          ' Close the serial port
          '
          If serialPort.IsOpen Then
            serialPort.Close()
          End If
      End Select

      '
      ' Reset Global Variables
      '
      gDevicePoweredOn = False
      gDeviceResponse = False
      gDeviceConnected = False
      gDeviceRefreshSystemState = False

      Return True

    Catch pEx As Exception
      '
      ' Process program exception
      '
      Call ProcessError(pEx, "DisconnectFromDevice()")
      Return False
    Finally
      '
      ' Update the Audio Device connection status
      '
      UpdateAudioConnectionDevice(Me.m_ConnectionType, Me.m_DeviceId, gDeviceConnected)
    End Try

  End Function

  ''' <summary>
  ''' Reconnect the connection to Audio Device
  ''' </summary>
  ''' <remarks></remarks>
  Public Sub Reconnect()

    '
    ' Ensure plug-in is disconnected
    '
    DisconnectFromDevice()

    '
    ' Ensure watchdog is not disabled
    '
    m_WatchdogDisabled = False

    '
    ' Ensure Device is marked as initialized
    '
    gDeviceInitialized = True

    '
    ' Interrupt the watchdog thread
    '
    If WatchdogThread.ThreadState = ThreadState.WaitSleepJoin Then
      If m_WatchdogActive = False Then
        WatchdogThread.Interrupt()
      End If
    End If

  End Sub

  ''' <summary>
  ''' isconnect from the Audio Device
  ''' </summary>
  ''' <remarks></remarks>
  Public Sub Disconnect()

    '
    ' Ensure the watchdog is disabled
    '
    m_WatchdogDisabled = True

    '
    ' Disconnect from the Device
    '
    DisconnectFromDevice()

  End Sub

#End Region

#Region "Audio Device Protocol Processing"

  ''' <summary>
  '''  Processes a received data string
  ''' </summary>
  ''' <param name="packet"></param>
  ''' <remarks></remarks>
  Private Sub ProcessReceived(ByVal packet As Byte())

    Dim strMessage As String = ""

    strMessage = String.Format("'{0}' bytes sent by Audio Device '{1}'.", packet.Length, m_ConnectionAddr)
    Call WriteMessage(strMessage, MessageType.Debug)

    '
    ' Update the last response date/time
    '
    m_DeviceLastResponse = DateTime.Now

    Try
      Dim bValidChecksum As Boolean = VerifyChecksum(packet)
      If bValidChecksum = False Then
        WriteMessage(String.Format("Audio Device packet does not contain a valid checksum."), MessageType.Warning)
        Return
      End If

      gDeviceResponse = True
      gDevicePoweredOn = True

      '
      ' Fix Packet
      '
      Dim message As Byte() = FixPacket(packet)

      ' =========================================
      ' Get the message header
      '
      Dim startOfMesasge As Integer = packet(0)     ' F0

      '
      ' Target Device Id
      ' 
      ' 
      Dim targetControllerId As Byte = message(1)   ' 0x00 is the system controller.  Values are 0x00=1 through 0x05=6 for a maximum of 6 controllers.  0x7F can be used for "all keypads", 0x7D = All Devices
      Dim targetZoneId As Byte = message(2)         ' 0x00 - 0x03 for 4 zone controllers and 0x00 - 0x05 for 6 zone controllers.  0x07D for peripheral devices (tuner, etc.)
      Dim targetKeypadId As Byte = message(3)       ' 7F is the controller itself, 7D targets all keypads on a particular zone for all sources for connected peripheral devices.  79 = Source Broadcast display message

      '
      ' Source Device Id
      '
      Dim sourceControllerId As Byte = message(4)   ' For 3rd party systems, use 0x00
      Dim sourceZoneId As Byte = message(5)         ' For 3rd party systems, use unique value
      Dim sourceKeypadId As Byte = message(6)       ' For 3rd party systems, use 0x07

      '
      ' Pakcet Type
      '
      Dim packetType As Byte = message(7)

      ' =========================================
      ' Process the message body
      '
      Select Case packetType
        Case PacketMessageType.SetData
          '
          ' Set data (set a parameter's value)
          '
          Dim targetPath As String = GetTargetPath(message)
          Dim sourcePath As String = GetSourcePath(message)

          If Regex.IsMatch(sourcePath, "5\.2\.0\.\d+\.0\.[0-7]") = True Then
            '
            ' This is a tone control SetData event
            '
            Dim zoneNumber As Byte = message(12)
            Dim controlType As Byte = message(14)
            Dim controlValue As Byte = message(21)

            Select Case controlType
              Case ZoneControlType.Bass             ' 5.2.0.zz.0.0
                SetAudioZoneValue(targetControllerId, zoneNumber, "audiozone-bass", controlValue)

              Case ZoneControlType.Treble           ' 5.2.0.zz.0.1
                SetAudioZoneValue(targetControllerId, zoneNumber, "audiozone-treble", controlValue)

              Case ZoneControlType.Loudness         ' 5.2.0.zz.0.2
                SetAudioZoneValue(targetControllerId, zoneNumber, "audiozone-loudness", controlValue)

              Case ZoneControlType.Balance          ' 5.2.0.zz.0.3
                SetAudioZoneValue(targetControllerId, zoneNumber, "audiozone-balance", controlValue)

              Case ZoneControlType.TurnOnVolume     ' 5.2.0.zz.0.4
                SetAudioZoneValue(targetControllerId, zoneNumber, "audiozone-turnon-vol", controlValue)

              Case ZoneControlType.BackgroundColor  ' 5.2.0.zz.0.5
                SetAudioZoneValue(targetControllerId, zoneNumber, "audiozone-keypad-bg-color", controlValue)

              Case ZoneControlType.DoNotDisturb     ' 5.2.0.zz.0.6
                SetAudioZoneValue(targetControllerId, zoneNumber, "audiozone-dnd", controlValue)

              Case ZoneControlType.PartyMode        ' 5.2.0.zz.0.7
                SetAudioZoneValue(targetControllerId, zoneNumber, "audiozone-partymode", controlValue)

            End Select

          ElseIf Regex.IsMatch(sourcePath, "4\.2\.0\.\d+\.[0-7]") = True Then

            Dim zoneNumber As Byte = message(12)
            Dim eventType As Byte = message(13)

            Select Case eventType
              Case PacketEventType.ZoneVolume
                Dim zoneVolume As Byte = message(20)

                SetAudioZoneValue(targetControllerId, zoneNumber, "audiozone-volume", zoneVolume)

              Case PacketEventType.ZoneSourceInput ' Return message from a Get Source request

                Dim zoneSource As Byte = message(20)

              Case PacketEventType.ZoneAllInfo ' 4.2.0.zz.7 Return message from a Get All Zone Info request

                Dim zonePower As Byte = message(20)
                ' 0x00 = OFF or 0x01 = ON
                SetAudioZoneValue(targetControllerId, zoneNumber, "audiozone-power", zonePower)

                Dim zoneSource As Byte = message(21)
                ' Current Input Selected -1
                SetAudioZoneValue(targetControllerId, zoneNumber, "audiozone-source", zoneSource)

                Dim zoneVolume As Byte = message(22)
                ' 0x00 through 0x32, 0x00 = 0 Displayed, 0x32 = 100 Displayed
                SetAudioZoneValue(targetControllerId, zoneNumber, "audiozone-volume", zoneVolume)

                Dim zoneBass As Byte = message(23)
                ' 0x00 = -10, 0x0A = Flat, 0x14 = +10
                SetAudioZoneValue(targetControllerId, zoneNumber, "audiozone-bass", zoneBass)

                Dim zoneTreble As Byte = message(24)
                ' 0x00 = -10, 0x0A = Flat, 0x14 = +10
                SetAudioZoneValue(targetControllerId, zoneNumber, "audiozone-treble", zoneTreble)

                Dim zoneLoudness As Byte = message(25)
                ' 0x00 = OFF or 0x01 = ON
                SetAudioZoneValue(targetControllerId, zoneNumber, "audiozone-loudness", zoneLoudness)

                Dim zoneBalance As Byte = message(26)
                ' 0x00 = More Left, 0x0A = Center, 0x14 = More Right
                SetAudioZoneValue(targetControllerId, zoneNumber, "audiozone-balance", zoneBalance)

                Dim systemOnState As Byte = message(27)
                ' 0x00 = All Zone Off or 0x01 = Any Zone is On
                SetAudioZoneValue(targetControllerId, zoneNumber, "audiozone-system-onstate", systemOnState)

                Dim sourceShared As Byte = message(28)
                ' 0x00 = Not Shared or 0x01 = Shared with Another Zone
                SetAudioZoneValue(targetControllerId, zoneNumber, "audiozone-source-shared", sourceShared)

                Dim partyMode As Byte = message(29)     ' ** Not for CAS44 or CAA66
                ' 0x00 = OFF or 0x01 = On, 0x02 = Master
                SetAudioZoneValue(targetControllerId, zoneNumber, "audiozone-partymode", partyMode)

                Dim DoNotDisturb As Byte = message(30)  ' ** Not for CAS44 or CAA66
                ' 0x00 = OFF or 0x01 = ON
                SetAudioZoneValue(targetControllerId, zoneNumber, "audiozone-dnd", DoNotDisturb)

              Case Else
                'Console.WriteLine(PrintPacket(message, True))
                'Console.WriteLine("Unknown event type {0}", eventType)

            End Select

          ElseIf targetPath = "2.1.0" And sourcePath = "2.1.0" Then
            '
            ' Zone On/off Event
            '
            Console.WriteLine("Known source path received {0}", sourcePath)

          ElseIf targetPath = "2.1.1" Then

            If targetKeypadId = &H79 Then
              '
              ' Process a Source Broadcast Display Message
              '
              Dim sourceNumber As Byte = sourceKeypadId
              Dim packetNumLoByte As Byte = message(14)
              Dim packetNumHiByte As Byte = message(15)
              Dim numPacketsLoByte As Byte = message(16)
              Dim numPacketsHiByte As Byte = message(17)
              Dim numDataBytesLoByte As Byte = message(18)
              Dim numDataBytesHiByte As Byte = message(19)
              Dim messageTypeandSourceNumber As Byte = message(20)
              Dim renderFlashTimeLo As Byte = message(21)     ' In 10ms increments. 0=permanent
              Dim renderFlashTimeHi As Byte = message(22)

              Dim zoneNumber As Byte = messageTypeandSourceNumber Or sourceKeypadId

              Dim renderFlashTime As Integer = ((renderFlashTimeHi << 8) Or (renderFlashTimeLo And &H00FF))
              If renderFlashTime <> 0 Then Exit Sub

              If numDataBytesLoByte > 0 Then
                Dim keypadMessage As String = System.Text.Encoding.ASCII.GetString(message, 23, numDataBytesLoByte - 4).TrimEnd(vbNullChar)
                SetKeypadSourceMessage(sourceNumber, keypadMessage)
              End If
            Else
              'Console.WriteLine("Unknown Targetpath/SourcePath received {0}/{1}", targetPath, sourcePath)
            End If

          ElseIf Regex.IsMatch(targetPath, "3\.4\.4\.[1-5]") = True And sourcePath = "2.4.0" Then
            '
            ' Handle Special Status / Display Element Messages
            '
            Dim state As Byte = message(21)
            Select Case targetPath
              Case "3.4.4.1"  ' Do Not Disturb Indication 

              Case "3.4.4.2"  ' Shared Indication

              Case "3.4.4.3"  ' System On Indication

              Case "3.4.4.4"  ' Party Mode Indication

              Case "3.4.4.5"  ' Party Master Indication

            End Select

          Else

            'Console.WriteLine("Unknown Targetpath/SourcePath received {0}/{1}", targetPath, sourcePath)

          End If

        Case PacketMessageType.RequestData
          '
          ' Request data (request a parameter's value)
          '
        Case PacketMessageType.Handshake
          '
          ' Handshake (acknowledges a data send)
          '
        Case PacketMessageType.Event        ' Event (triggers a system response that may set a paramter value, update displays, etc...)
          Dim targetPath As String = GetTargetPath(message)
          Dim sourcePath As String = GetSourcePath(message)

          If targetKeypadId = &H7D Then

          ElseIf sourceKeypadId = &H7D Then

          ElseIf targetKeypadId = &H7F Then

            If sourceZoneId < 125 Then
              gDelayCounter = 1
              GetZoneStatusFromController(sourceControllerId, sourceZoneId)
            End If

          ElseIf sourceKeypadId = &H7F Then

          End If

        Case PacketMessageType.Display
          '
          ' Process Locally Rendered Display Messages
          '   Rendering requires the display device to have a table of rendering functions, a master string table and string ID tables associated with 
          '   many of the rendering types. A render display message includes a rendering type and the value to be rendered.
          '   For render types such as volume, bass, treble, etc the rendering will use the value in the rendered display. 
          '   (E.g. data  = 20. display = "Volume: 20").
          Dim renderValueLo As Byte = message(8)
          Dim renderValueHi As Byte = message(9)
          Dim renderFlashTimeLo As Byte = message(10)     ' In 10ms increments. 0=permanent
          Dim renderFlashTimeHi As Byte = message(11)
          Dim renderType As Byte = message(12)

          Dim renderFlashTime As Integer = ((renderFlashTimeHi << 8) Or (renderFlashTimeLo And &H00FF))
          If renderFlashTime <> 0 Then Exit Sub

          Select Case renderType
            Case 5
              ' RenderType_SOURCE_NAME
              ' renderValueLo = DISPLAYSTRINGS_sourceNames
              ' renderValueHi = Used to indicate selected source
              Dim keypadMessage As String = GetSourceName(renderValueLo)
              SetKeypadMessage(targetControllerId, sourceZoneId, keypadMessage)

            Case 16
              ' RenderType_VOLUME
              Dim keypadMessage As String = GetZoneVolumeLevel(renderValueLo)
              SetKeypadMessage(targetControllerId, sourceZoneId, keypadMessage)

            Case 17
              ' RenderType_BASS
              Dim keypadMessage As String = GetZoneBassLevel(renderValueLo)
              SetKeypadMessage(targetControllerId, sourceZoneId, keypadMessage)

            Case 18
              ' RenderType_TREBLE
              Dim keypadMessage As String = GetZoneTrebleLevel(renderValueLo)
              SetKeypadMessage(targetControllerId, sourceZoneId, keypadMessage)

            Case 19
              ' RenderType_BALANCE
              Dim keypadMessage As String = GetZoneBalanceLevel(renderValueLo)
              SetKeypadMessage(targetControllerId, sourceZoneId, keypadMessage)

            Case 24
              ' RenderType_ALL_STRINGS 24 
              Dim keypadMessage As String = GetRussoundMasterString(renderValueLo)
              SetKeypadMessage(targetControllerId, sourceZoneId, keypadMessage)

            Case Else
              'Console.WriteLine("Received Render Type Unknown:  Controller {0}, Zone {1}, Volume {2}", targetControllerId, sourceZoneId, renderValueLo.ToString("X2"))

          End Select

        Case Else
          'Console.WriteLine("Unknown message type {0}", packetType)

      End Select

      '
      ' End of Packet is F7
      '

    Catch pEx As Exception
      WriteMessage(pEx.Message, MessageType.Error)
    End Try

  End Sub

  ''' <summary>
  ''' Calculate the Target Path
  ''' </summary>
  ''' <param name="message"></param>
  ''' <returns></returns>
  Private Function GetTargetPath(ByRef message() As Byte) As String

    Dim path As New StringBuilder

    Try
      '
      ' Calculate Target Path
      '
      Dim targetPathByte As Integer = 8
      Dim targetPathBytes As Byte = message(targetPathByte)
      Dim targetPath As String = String.Empty

      path.Append(targetPathBytes.ToString)

      For i = 1 To targetPathBytes
        path.Append(".")
        path.Append(message(targetPathByte + i).ToString)
      Next

    Catch pEx As Exception
      Return "0"
    End Try

    Return path.ToString

  End Function

  ''' <summary>
  ''' Calculate the Source Path
  ''' </summary>
  ''' <param name="message"></param>
  ''' <returns></returns>
  Private Function GetSourcePath(ByRef message() As Byte) As String

    Dim path As New StringBuilder

    Try
      '
      ' Calculate Source Path
      '
      Dim targetPathByte As Integer = 8
      Dim targetPathBytes As Byte = message(targetPathByte)
      Dim sourcePathByte As Integer = targetPathByte + 1 + targetPathBytes
      Dim sourcePathBytes As Byte = message(sourcePathByte)

      path.Append(sourcePathBytes.ToString)

      For i = 1 To sourcePathBytes
        path.Append(".")
        path.Append(message(sourcePathByte + i).ToString)
      Next

    Catch pEx As Exception
      Return "0"
    End Try

    Return path.ToString

  End Function

  ''' <summary>
  ''' Formats packet for sending to Audio Controller
  ''' </summary>
  ''' <param name="packet"></param>
  ''' <returns></returns>
  Private Function FormatPacket(ByRef packet As Byte()) As Byte()

    Dim message(0) As Byte

    Try

      Dim x As Integer = 0

      For i = 0 To packet.Length - 1

        If packet(i) > &HF1 AndAlso i < packet.Length - 2 Then
          'Console.WriteLine("Insert Special Invert Character added in packet {0}", i.ToString)

          ReDim Preserve message(x)
          message(x) = &HF1

          ' Skip to next packet
          x += 1
          ReDim Preserve message(x)
          message(x) = packet(i) Xor 255
        Else

          ReDim Preserve message(x)
          message(x) = packet(i)

        End If

        x += 1
      Next

    Catch pEx As Exception
      WriteMessage(pEx.Message, MessageType.Error)
    End Try

    Return message

  End Function

  ''' <summary>
  ''' Adjusts incoming packet
  ''' </summary>
  ''' <param name="packet"></param>
  ''' <returns></returns>
  Private Function FixPacket(ByRef packet As Byte()) As Byte()

    Dim message(0) As Byte

    Try

      Dim x As Integer = 0

      For i = 0 To packet.Length - 1
        ReDim Preserve message(x)

        If packet(i) = &HF1 Then
          'Console.WriteLine("Special Invert Character detected in packet {0}", i.ToString)
          ' Skip to next packet
          i += 1
          message(x) = packet(i) Xor 255
        Else
          message(x) = packet(i)
        End If

        x += 1
      Next

    Catch pEx As Exception
      WriteMessage(pEx.Message, MessageType.Error)
    End Try

    Return message

  End Function

  ''' <summary>
  ''' Formats packet for printing
  ''' </summary>
  ''' <param name="packet"></param>
  ''' <returns></returns>
  Private Function PrintPacket(ByVal packet As Byte(), Optional bHeader As Boolean = True) As String

    Dim sb As New StringBuilder

    Try

      If bHeader = True Then
        For i = 1 To packet.Length
          sb.Append(String.Format("{0}|", i.ToString.PadLeft(2, "0")))
        Next
        sb.AppendLine()
      End If

      For i = 0 To packet.Length - 1
        sb.Append(String.Format("{0:X2}|", packet(i)))
      Next
      sb.AppendLine()

    Catch pEx As Exception

    End Try

    Return sb.ToString

  End Function

  ''' <summary>
  ''' Verify packet checksum
  ''' </summary>
  ''' <param name="packet"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Function VerifyChecksum(ByRef packet() As Byte) As Boolean

    Try

      '
      ' Get Checksum from packet
      '
      Dim pktChecksum As Byte = packet(packet.Length - 2)

      ' Step 1
      Dim byteCount As Integer = packet.Length - 2
      Dim Checksum As Integer = 0
      For i As Integer = 0 To byteCount - 1
        Checksum += packet(i)
      Next

      ' Step 2
      Checksum += byteCount

      ' Step 3
      Dim pktChecksum2 As Byte = Checksum And 127

      ' Return Results
      Return pktChecksum = pktChecksum2

    Catch pEx As Exception
      Return False
    End Try

  End Function

  ''' <summary>
  ''' Calculate packet checksum
  ''' </summary>
  ''' <param name="packet"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Function CalculateChecksum(ByRef packet() As Byte) As Byte

    Try
      Dim byteCount As Integer = packet.Length

      ' Step 1
      Dim Checksum As Integer = 0
      For i As Integer = 0 To byteCount - 2
        Checksum += packet(i)
      Next

      ' Step 2
      Checksum += byteCount - 2

      ' Step 3
      Dim pktChecksum As Byte = Checksum And 127

      ' Return Checksum
      Return pktChecksum

    Catch pEx As Exception
      Return 0
    End Try

  End Function

  ''' <summary>
  ''' Sends command to Audio Device
  ''' </summary>
  ''' <param name="packet"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function SendToDevice(ByVal packet() As Byte) As Boolean

    Dim strMessage As String = ""

    strMessage = "Entered SendToDevice() function."
    Call WriteMessage(strMessage, MessageType.Debug)

    Try

      Select Case m_ConnectionType.ToUpper
        Case "ETHERNET"
          '
          ' Set data to Ethernet connection
          '
          If gDeviceConnected = True And gIOEnabled = True Then
            strMessage = String.Format("Sending '{0}' bytes to Audio Device device '{1}' via Ethernet.", packet.Count, m_ConnectionAddr)
            Call WriteMessage(strMessage, MessageType.Debug)
            Return SendMessageToEthernet(packet)
          Else
            Return False
          End If

        Case "SERIAL"
          '
          ' Send data using the serial port
          '
          If serialPort.IsOpen = True Then
            strMessage = String.Format("Sending '{0}' bytes to Audio Device device '{1}' via Serial.", packet.Count, m_ConnectionAddr)
            Call WriteMessage(strMessage, MessageType.Debug)

            serialPort.Write(packet, 0, packet.Count)
          End If
          Return serialPort.IsOpen

        Case Else
          Return False

      End Select

    Catch pEx As Exception
      '
      ' Process program exception
      '
      Call ProcessError(pEx, "SendToDevice()")
      Return False
    End Try

  End Function

#End Region

#Region "Audio Device Connection Status"

  ''' <summary>
  ''' Checks to see if we are connected to the Audio Device
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function CheckDeviceConnection() As Boolean

    Dim dtStartTime As Date
    Dim etElapsedTime As TimeSpan
    Dim iMillisecondsWaited As Integer
    Dim Attempts As Integer = 0

    Try
      '
      ' We are going to use this function as a verification test to see if the Audio Device is actually connected
      '
      Call WriteMessage(String.Format("Checking if the Audio Device device '{0}' is connected ...", m_ConnectionAddr), MessageType.Debug)

      If gDeviceConnected = False Then
        Return False
      End If

      '
      ' Reset global variable
      '
      gDeviceResponse = False

      Do
        '
        ' Get the power status
        '
        AddCommand(GetZonePowerState(0, 0))

        '
        ' Block for until we get our Audio Device response
        '
        dtStartTime = Now
        Do
          Thread.Sleep(50)
          etElapsedTime = Now.Subtract(dtStartTime)
          iMillisecondsWaited = etElapsedTime.Milliseconds + (etElapsedTime.Seconds * 1000)
        Loop While gDeviceResponse = False And iMillisecondsWaited < 3000

        If gDeviceResponse = False Then
          Attempts = Attempts + 1
        End If

      Loop While Attempts < 3 And gDeviceResponse = False

      Return gDeviceResponse

    Catch pEx As Exception
      '
      ' Process program exception
      '
      Call ProcessError(pEx, "CheckDeviceConnection()")
      Return False
    End Try

  End Function

  ''' <summary>
  ''' Checks to see if we are connected to the Audio Device
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function CheckPowerStatus() As Boolean

    Dim dtStartTime As Date
    Dim etElapsedTime As TimeSpan
    Dim iMillisecondsWaited As Integer
    Dim Attempts As Integer = 0

    Try

      If gDeviceConnected = False Then
        Return False
      End If

      '
      ' We are going to use this function as a verification test to see if the Audio Device is actually connected
      '
      Call WriteMessage(String.Format("Checking if the Audio Device device '{0}' is powered on ...", m_ConnectionAddr), MessageType.Debug)

      '
      ' Reset Global Variable 
      '
      gDeviceResponse = False

      Do
        '
        ' Get the power status
        '
        AddCommand(GetZonePowerState(0, 0))

        '
        ' Block for until we get our Audio Device response
        '
        dtStartTime = Now
        Do
          Thread.Sleep(50)
          etElapsedTime = Now.Subtract(dtStartTime)
          iMillisecondsWaited = etElapsedTime.Milliseconds + (etElapsedTime.Seconds * 1000)
        Loop While gDeviceResponse = False And iMillisecondsWaited < 3000

        If gDeviceResponse = False Then
          Attempts = Attempts + 1
        End If

      Loop While Attempts < 3 And gDeviceResponse = False

      '
      ' Process the results and return the status
      '
      If gDeviceResponse = False Then
        WriteMessage(String.Format("The Audio Device device '{0}' did not respond.", m_ConnectionAddr), MessageType.Debug)
        Return False
      ElseIf gDevicePoweredOn = False Then
        WriteMessage(String.Format("The Audio Device device '{0}' is powered off.", m_ConnectionAddr), MessageType.Debug)
        Return False
      Else
        WriteMessage(String.Format("The Audio Device device '{0}' is powered on.", m_ConnectionAddr), MessageType.Debug)
        Return True
      End If

    Catch pEx As Exception
      '
      ' Process program exception
      '
      Call ProcessError(pEx, "CheckPowerStatus()")
      Return False
    End Try

  End Function

  ''' <summary>
  ''' Determine if Device connection is active
  ''' </summary>
  ''' <remarks></remarks>
  Public Sub CheckPhysicalConnection()

    Try

      Select Case m_ConnectionType.ToUpper
        Case "ETHERNET"
          If TcpClient.Connected = False Or TcpClient.Client.Connected = False Then
            gDeviceConnected = False
          Else

            If NetworkStream.CanRead = False Or NetworkStream.CanWrite = False Then
              gDeviceConnected = False
            End If

          End If
        Case "SERIAL"
          gDeviceConnected = serialPort.IsOpen
      End Select

    Catch pEx As Exception

    End Try

  End Sub

  ''' <summary>
  ''' Fail-safe block
  ''' </summary>
  ''' <param name="iSecondsToWait"></param>
  ''' <remarks></remarks>
  Private Sub FailSafeBlock(ByVal iSecondsToWait As Integer)

    Dim dtStartTime As Date
    Dim etElapsedTime As TimeSpan

    Dim iMillisecondsWaited As Integer
    Dim iMilliSecondsToWait As Integer = iSecondsToWait * 1000

    '
    ' Block until all the previous commands have been processed - always implement a block fail-safe
    '
    dtStartTime = Now
    Do
      Thread.Sleep(50)
      etElapsedTime = Now.Subtract(dtStartTime)
      iMillisecondsWaited = etElapsedTime.Milliseconds + (etElapsedTime.Seconds * 1000)
    Loop While gCommandQueue.Count() > 0 And iMillisecondsWaited < iMilliSecondsToWait

  End Sub

#End Region

#Region "Audio Device Command Processing"

  ''' <summary>
  ''' Adds command to command buffer for processing
  ''' </summary>
  ''' <param name="Packet"></param>
  ''' <param name="bForce"></param>
  ''' <remarks></remarks>
  Protected Sub AddCommand(ByVal Packet As Byte(), Optional ByVal bForce As Boolean = False)

    Try
      '
      ' bForce may be used to add a command to repeat a command
      '
      If gDeviceConnected = True Then
        Dim bCommandInQueue As Boolean = False

        SyncLock gCommandQueue.SyncRoot
          '
          ' Check Command Queue
          '
          For Each Buffer As Byte() In gCommandQueue
            If Packet.SequenceEqual(Buffer) = True Then
              bCommandInQueue = True
              Exit For
            End If
          Next

          If bCommandInQueue = False Or bForce = True Then
            gCommandQueue.Enqueue(Packet)
          Else
            WriteMessage(String.Format("Ignoring command because it is already in the queue for Device device '{0}'.", m_ConnectionAddr), MessageType.Debug)
          End If

        End SyncLock

      Else
        WriteMessage(String.Format("Ignoring command because the Audio Device device '{0}' is not connected.", m_ConnectionAddr), MessageType.Debug)
      End If

    Catch pEx As Exception
      '
      ' Process program exception
      '
      Call ProcessError(pEx, "AddCommand()")
    End Try

  End Sub

  ''' <summary>
  ''' Process Commands in the Command Queue
  ''' </summary>
  Public Sub ProcessCommandQueue()

    Dim bAbortThread As Boolean = False

    Try

      While bAbortThread = False

        '
        ' Don't send a command if the plug-in is not initialized
        '
        If gCommandQueue.Count > 0 AndAlso gIOEnabled = False Then

          If gIOEnabled = False Then

            Dim CommandCount As Integer = gCommandQueue.Count

            Dim strMessage As String = String.Format("The {0} plug-in is no longer properly initialized.  Flushing {1} commands from the command queue.", IFACE_NAME, CommandCount)
            Call WriteMessage(strMessage, MessageType.Warning)

            '
            ' Clear out all commands from the queue
            '
            SyncLock gCommandQueue.SyncRoot
              gCommandQueue.Clear()
            End SyncLock

            '
            ' Sleep so we don't fill up the logs
            '
            Thread.Sleep(1000 * 30)
          End If

        Else

          While gCommandQueue.Count > 0 AndAlso gIOEnabled = True

            While gDelayCounter > 0
              gDelayCounter -= 1
              Thread.Sleep(1000)
            End While

            '
            ' Process the next command in the queue
            '
            SyncLock gCommandQueue.SyncRoot
              Dim Buffer As Byte() = gCommandQueue.Peek
              SendToDevice(Buffer)
            End SyncLock

            '
            ' We don't care if we get a response
            '
            SyncLock gCommandQueue.SyncRoot
              gCommandQueue.Dequeue()
            End SyncLock

            Thread.Sleep(200)
          End While

        End If

        '
        ' Give up some time
        '
        Thread.Sleep(25)

      End While ' Stay in thread until we get an abort/exit request

    Catch pEx As ThreadAbortException
      ' 
      ' There was a normal request to terminate the thread.  
      '
      bAbortThread = True      ' Not actually needed
      WriteMessage(String.Format("ProcessCommandQueue thread received abort request, terminating normally."), MessageType.Informational)

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "ProcessCommandQueue()")

    Finally
      '
      ' Notify that we are exiting the thread 
      '
      WriteMessage(String.Format("ProcessCommandQueue terminated."), MessageType.Debug)

    End Try

  End Sub

#End Region

#Region "Audio Device Commands"

  ''' <summary>
  ''' Processes commands and waits for the response
  ''' </summary>
  ''' <remarks></remarks>
  Protected Sub RefreshSystemState()

    Dim bAbortThread As Boolean = False

    Try

      While bAbortThread = False
        '
        ' Process commands in command queue
        '
        While gDeviceConnected = True And gIOEnabled = True

          If gDeviceRefreshSystemState = True Then
            gDeviceRefreshSystemState = False
            '
            ' Get the Zone Status for each controller and zone
            '
            RefreshZoneStatus()
          End If

          Thread.Sleep(1000)
        End While

        '
        ' Give up some time
        '
        Thread.Sleep(1000)

      End While ' Stay in thread until we get an abort/exit request

    Catch pEx As ThreadAbortException
      ' 
      ' There was a normal request to terminate the thread.  
      '
      bAbortThread = True      ' Not actually needed
      WriteMessage(String.Format("RefreshSystemState thread received abort request, terminating normally."), MessageType.Debug)

    Catch pEx As Exception
      '
      ' Return message
      '
      ProcessError(pEx, "RefreshSystemState()")

    Finally
      '
      ' Notify that we are exiting the thread 
      '
      WriteMessage(String.Format("RefreshSystemState terminated."), MessageType.Debug)
    End Try

  End Sub

  ''' <summary>
  ''' Refreshses the Zone Status for All Zones
  ''' </summary>
  Private Sub RefreshZoneStatus()

    Try

      Dim zones As Integer = 0
      For cc = 0 To 5
        For zz = 0 To 5
          zones += 1
          If zones <= m_DeviceZones Then
            GetZoneStatusFromController(cc, zz)
          End If
        Next
      Next

      '
      ' Wait 120 seconds
      '
      FailSafeBlock(120)

    Catch pEx As Exception

    End Try

  End Sub

  ''' <summary>
  ''' Gets all values for the controller/zone
  ''' </summary>
  ''' <param name="cc"></param>
  ''' <param name="zz"></param>
  Private Sub GetZoneStatusFromController(ByVal cc As Byte, ByVal zz As Byte)

    Try

      AddCommand(GetZoneTurnOnVolume(cc, zz))
      AddCommand(GetZoneBackgroundColor(cc, zz))
      AddCommand(GetZoneStatus(cc, zz))

    Catch pEx As Exception

    End Try

  End Sub

#Region "Zone Power State"

  ''' <summary>
  ''' Set Zone Power State Ex
  ''' </summary>
  ''' <param name="cc"></param>
  ''' <param name="zz"></param>
  ''' <param name="state"></param>
  Public Sub SetZoneStateEx(ByVal cc As Byte, zz As Byte, state As Byte)
    AddCommand(SetZoneState(cc, zz, state))
    AddCommand(GetZoneStatus(cc, zz))
  End Sub

  ''' <summary>
  ''' Get Zone Power State
  ''' </summary>
  ''' <param name="cc"></param>
  ''' <param name="zz"></param>
  ''' <returns></returns>
  Protected Function GetZonePowerState(ByVal cc As Byte, zz As Byte) As Byte()

    Dim buffer(16) As Byte

    Try
      Dim kk As Byte = IIf(m_DeviceModel = "ACA-E5", &H71, &H70)

      buffer.SetValue(CByte(&HF0), 0)       ' F0 Start of Message
      buffer.SetValue(CByte(cc), 1)         ' Target Controller Id
      buffer.SetValue(CByte(&H00), 2)       ' Target Zone Id
      buffer.SetValue(CByte(&H7F), 3)       ' Target Keypad Number 

      buffer.SetValue(CByte(cc), 4)         ' Source Controller Id
      buffer.SetValue(CByte(zz), 5)         ' Source Zone Id
      buffer.SetValue(CByte(kk), 6)         ' Source Keypad Number

      buffer.SetValue(PacketMessageType.RequestData, 7) ' Message Type (1 = Set)

      buffer.SetValue(CByte(&H04), 8)       ' Target Path:  0x04 (Represented by the following 4 bytes)
      buffer.SetValue(CByte(&H02), 9)       ' Path Level 0: 0x02 (Root Menu)
      buffer.SetValue(CByte(&H00), 10)      ' Path Level 1: 0x00 (Run Menu)
      buffer.SetValue(CByte(zz), 11)        ' Path Level 2: zz   (Zone number -1)
      buffer.SetValue(CByte(&H06), 12)      ' Path Level 3: 0x06 (Power)
      buffer.SetValue(CByte(&H00), 13)      ' Source Path:  0x00 (this packet does not contain a source path)

      buffer.SetValue(CByte(&H00), 14)      ' 0x00

      Dim checkSum As Byte = CalculateChecksum(buffer)
      buffer.SetValue(CByte(checkSum), 15)
      buffer.SetValue(CByte(&HF7), 16)

    Catch pEx As Exception

    End Try

    Return buffer

  End Function

  ''' <summary>
  ''' Set Zone Power State
  ''' </summary>
  ''' <param name="cc"></param>
  ''' <param name="zz"></param>
  ''' <param name="state"></param>
  ''' <returns></returns>
  Protected Function SetZoneState(ByVal cc As Byte, zz As Byte, state As Byte) As Byte()

    Dim buffer(21) As Byte

    Try

      Dim kk As Byte = IIf(m_DeviceModel = "ACA-E5", &H71, &H70)

      buffer.SetValue(CByte(&HF0), 0)       ' F0 Start of Message
      buffer.SetValue(CByte(cc), 1)         ' Target Controller Id
      buffer.SetValue(CByte(&H00), 2)       ' Target Zone Id
      buffer.SetValue(CByte(&H7F), 3)       ' Target Keypad Number 

      buffer.SetValue(CByte(cc), 4)         ' Source Controller Id
      buffer.SetValue(CByte(zz), 5)         ' Source Zone Id
      buffer.SetValue(CByte(kk), 6)         ' Source Keypad Number (0x70 or 0X71 when connected to a ACA-E5)

      buffer.SetValue(PacketMessageType.Event, 7) ' Message Type (5 = Event)

      buffer.SetValue(CByte(&H02), 8)
      buffer.SetValue(CByte(&H02), 9)
      buffer.SetValue(CByte(&H00), 10)
      buffer.SetValue(CByte(&H00), 11)
      buffer.SetValue(CByte(&HF1), 12)
      buffer.SetValue(CByte(&H23), 13)
      buffer.SetValue(CByte(&H00), 14)
      buffer.SetValue(CByte(state), 15)     ' Desired Power State
      buffer.SetValue(CByte(&H00), 16)
      buffer.SetValue(CByte(zz), 17)        ' Zone Id -1
      buffer.SetValue(CByte(&H00), 18)
      buffer.SetValue(CByte(&H01), 19)

      Dim checkSum As Byte = CalculateChecksum(buffer)
      buffer.SetValue(CByte(checkSum), 20)
      buffer.SetValue(CByte(&HF7), 21)

    Catch pEx As Exception

    End Try

    Return buffer

  End Function

  ''' <summary>
  ''' Set All Zone Power State
  ''' </summary>
  ''' <param name="cc"></param>
  ''' <param name="zz"></param>
  ''' <param name="state"></param>
  ''' <returns></returns>
  Protected Function SetZoneStateAll(ByVal cc As Byte, zz As Byte, state As Byte) As Byte()

    Dim buffer(21) As Byte

    Try

      Dim kk As Byte = IIf(m_DeviceModel = "ACA-E5", &H71, &H70)

      buffer.SetValue(CByte(&HF0), 0)       ' F0 Start of Message
      buffer.SetValue(CByte(&H7E), 1)       ' Target Controller Id (7E is the controller)
      buffer.SetValue(CByte(&H00), 2)       ' Target Zone Id
      buffer.SetValue(CByte(&H7F), 3)       ' Target Keypad Number 

      buffer.SetValue(CByte(cc), 4)         ' Source Controller Id
      buffer.SetValue(CByte(zz), 5)         ' Source Zone Id
      buffer.SetValue(CByte(kk), 6)         ' Source Keypad Number (0x70 or 0X71 when connected to a ACA-E5)

      buffer.SetValue(PacketMessageType.Event, 7) ' Message Type (5 = Event)

      buffer.SetValue(CByte(&H02), 8)
      buffer.SetValue(CByte(&H02), 9)
      buffer.SetValue(CByte(&H00), 10)
      buffer.SetValue(CByte(&H00), 11)
      buffer.SetValue(CByte(&HF1), 12)
      buffer.SetValue(CByte(&H23), 13)
      buffer.SetValue(CByte(&H00), 14)
      buffer.SetValue(CByte(state), 15)     ' Desired Power State
      buffer.SetValue(CByte(&H00), 16)
      buffer.SetValue(CByte(zz), 17)        ' Zone Id -1
      buffer.SetValue(CByte(&H00), 18)
      buffer.SetValue(CByte(&H01), 19)

      Dim checkSum As Byte = CalculateChecksum(buffer)
      buffer.SetValue(CByte(checkSum), 20)
      buffer.SetValue(CByte(&HF7), 21)

    Catch pEx As Exception

    End Try

    Return buffer

  End Function

#End Region

#Region "Zone Source"

  ''' <summary>
  ''' Set Zone Source State Ex
  ''' </summary>
  ''' <param name="cc"></param>
  ''' <param name="zz"></param>
  ''' <param name="source"></param>
  Public Sub SetZoneSourceEx(ByVal cc As Byte, ByVal zz As Byte, ByVal source As Byte)
    AddCommand(SetZoneSource(cc, zz, source))
    AddCommand(GetZoneStatus(cc, zz))
  End Sub

  ''' <summary>
  ''' Get Zone Power State
  ''' </summary>
  ''' <param name="cc"></param>
  ''' <param name="zz"></param>
  ''' <returns></returns>
  Protected Function GetZoneSource(ByVal cc As Byte, ByVal zz As Byte) As Byte()

    Dim buffer(16) As Byte

    Try
      Dim kk As Byte = IIf(m_DeviceModel = "ACA-E5", &H71, &H70)

      buffer.SetValue(CByte(&HF0), 0)       ' F0 Start of Message
      buffer.SetValue(CByte(cc), 1)         ' Target Controller Id
      buffer.SetValue(CByte(&H00), 2)       ' Target Zone Id
      buffer.SetValue(CByte(&H7F), 3)       ' Target Keypad Number

      buffer.SetValue(CByte(cc), 4)         ' Source Controller Id
      buffer.SetValue(CByte(zz), 5)         ' Source Zone Id
      buffer.SetValue(CByte(kk), 6)         ' Source Keypad Number

      buffer.SetValue(PacketMessageType.RequestData, 7) ' Message Type

      buffer.SetValue(CByte(&H04), 8)       ' Target Path:    0x04 (Represented by the following 4 bytes)
      buffer.SetValue(CByte(&H02), 9)       ' Path Level 0:   0x02 (Root Menu)
      buffer.SetValue(CByte(&H00), 10)      ' Path Level 1:   0x00 (Run Menu)
      buffer.SetValue(CByte(zz), 11)        ' Path Level 2:   zz   (Zone number -1)
      buffer.SetValue(CByte(&H02), 12)      ' Path Level 3:   0x02 (Source)
      buffer.SetValue(CByte(&H00), 13)      ' Source Path:    0x00 (this packet does not contain a source path)

      buffer.SetValue(CByte(&H00), 14)      ' 0x00

      Dim checkSum As Byte = CalculateChecksum(buffer)
      buffer.SetValue(CByte(checkSum), 15)
      buffer.SetValue(CByte(&HF7), 16)

    Catch pEx As Exception

    End Try

    Return buffer

  End Function

  ''' <summary>
  ''' Set Zone Source State
  ''' </summary>
  ''' <param name="cc"></param>
  ''' <param name="zz"></param>
  ''' <param name="source"></param>
  ''' <returns></returns>
  Protected Function SetZoneSource(ByVal cc As Byte, ByVal zz As Byte, ByVal source As Byte) As Byte()

    Dim buffer(21) As Byte

    Try
      Dim kk As Byte = IIf(m_DeviceModel = "ACA-E5", &H71, &H70)

      buffer.SetValue(CByte(&HF0), 0)       ' F0 Start of Message
      buffer.SetValue(CByte(cc), 1)         ' Target Controller Id
      buffer.SetValue(CByte(&H00), 2)       ' Target Zone Id
      buffer.SetValue(CByte(&H7F), 3)       ' Target Keypad Number

      buffer.SetValue(CByte(cc), 4)         ' Source Controller Id
      buffer.SetValue(CByte(zz), 5)         ' Source Zone Id
      buffer.SetValue(CByte(kk), 6)         ' Source Keypad Number

      buffer.SetValue(PacketMessageType.Event, 7)   ' Message Type (5 = Event)

      buffer.SetValue(CByte(&H02), 8)       ' 0x02
      buffer.SetValue(CByte(&H00), 9)       ' 0x00
      buffer.SetValue(CByte(&H00), 10)      ' 0x00
      buffer.SetValue(CByte(&H00), 11)      ' 0x00
      buffer.SetValue(CByte(&HF1), 12)      ' 0xF1
      buffer.SetValue(CByte(&H3E), 13)      ' 0x3E
      buffer.SetValue(CByte(&H00), 14)      ' 0x00
      buffer.SetValue(CByte(&H00), 15)      ' 0x00
      buffer.SetValue(CByte(&H00), 16)      ' 0x00
      buffer.SetValue(CByte(source), 17)    ' 0x00
      buffer.SetValue(CByte(&H00), 18)      ' 0x00
      buffer.SetValue(CByte(&H01), 19)      ' 0x01

      Dim checkSum As Byte = CalculateChecksum(buffer)
      buffer.SetValue(CByte(checkSum), 20)
      buffer.SetValue(CByte(&HF7), 21)

    Catch pEx As Exception

    End Try

    Return buffer

  End Function

#End Region

#Region "Zone Volume"
  ''' <summary>
  ''' Set Zone Turn On Volume
  ''' </summary>
  ''' <param name="cc"></param>
  ''' <param name="zz"></param>
  ''' <param name="volume"></param>
  Public Sub SetZoneVolumeEx(ByVal cc As Byte, ByVal zz As Byte, ByVal volume As Byte)
    AddCommand(SetZoneVolume(cc, zz, volume))
    AddCommand(GetZoneVolume(cc, zz))
  End Sub

  ''' <summary>
  ''' Get Zone Volume
  ''' </summary>
  ''' <param name="cc"></param>
  ''' <param name="zz"></param>
  ''' <returns></returns>
  Protected Function GetZoneVolume(ByVal cc As Byte, ByVal zz As Byte) As Byte()

    Dim buffer(16) As Byte

    Try
      Dim kk As Byte = IIf(m_DeviceModel = "ACA-E5", &H71, &H70)

      buffer.SetValue(CByte(&HF0), 0)       ' F0 Start of Message
      buffer.SetValue(CByte(cc), 1)         ' Target Controller Id
      buffer.SetValue(CByte(&H00), 2)       ' Target Zone Id
      buffer.SetValue(CByte(&H7F), 3)       ' Target Keypad Number

      buffer.SetValue(CByte(cc), 4)         ' Source Controller Id
      buffer.SetValue(CByte(zz), 5)         ' Source Zone Id
      buffer.SetValue(CByte(kk), 6)         ' Source Keypad Number

      buffer.SetValue(PacketMessageType.RequestData, 7) ' Message Type

      buffer.SetValue(CByte(&H04), 8)       ' Target Path:    0x04 (Represented by the following 4 bytes)
      buffer.SetValue(CByte(&H02), 9)       ' Path Level 0:   0x02 (Root Menu)
      buffer.SetValue(CByte(&H00), 10)      ' Path Level 1:   0x00 (Run Menu)
      buffer.SetValue(CByte(zz), 11)        ' Path Level 2:   zz   (Zone number -1)
      buffer.SetValue(CByte(&H01), 12)      ' Path Level 3:   0x01 (Volume)
      buffer.SetValue(CByte(&H00), 13)      ' Source Path:    0x00 (this packet does not contain a source path)

      buffer.SetValue(CByte(&H00), 14)      ' 0x00

      Dim checkSum As Byte = CalculateChecksum(buffer)
      buffer.SetValue(CByte(checkSum), 15)
      buffer.SetValue(CByte(&HF7), 16)

    Catch pEx As Exception

    End Try

    Return buffer

  End Function

  ''' <summary>
  ''' Set Zone Volume
  ''' </summary>
  ''' <param name="cc"></param>
  ''' <param name="zz"></param>
  ''' <param name="volume"></param>
  ''' <returns></returns>
  Protected Function SetZoneVolume(ByVal cc As Byte, ByVal zz As Byte, ByVal volume As Byte) As Byte()

    Dim buffer(21) As Byte

    Try
      Dim kk As Byte = IIf(m_DeviceModel = "ACA-E5", &H71, &H70)

      buffer.SetValue(CByte(&HF0), 0)       ' F0 Start of Message
      buffer.SetValue(CByte(cc), 1)         ' Target Controller Id
      buffer.SetValue(CByte(&H00), 2)       ' Target Zone Id
      buffer.SetValue(CByte(&H7F), 3)       ' Target Keypad Number

      buffer.SetValue(CByte(cc), 4)         ' Source Controller Id
      buffer.SetValue(CByte(zz), 5)         ' Source Zone Id
      buffer.SetValue(CByte(kk), 6)         ' Source Keypad Number

      buffer.SetValue(PacketMessageType.Event, 7) ' Message Type (5 = Event)

      buffer.SetValue(CByte(&H02), 8)       ' 0x02
      buffer.SetValue(CByte(&H02), 9)       ' 0x02
      buffer.SetValue(CByte(&H00), 10)      ' 0x00
      buffer.SetValue(CByte(&H00), 11)      ' 0x00

      buffer.SetValue(CByte(&HF1), 12)      ' 0xF1
      buffer.SetValue(CByte(&H21), 13)      ' 0x21
      buffer.SetValue(CByte(&H00), 14)      ' 0x00
      buffer.SetValue(volume, 15)           ' Volume Level 
      buffer.SetValue(CByte(&H00), 16)      ' 0x00

      buffer.SetValue(CByte(zz), 17)        ' Zone Number -1
      buffer.SetValue(CByte(&H00), 18)      ' 0x00
      buffer.SetValue(CByte(&H01), 19)      ' 0x01

      Dim checkSum As Byte = CalculateChecksum(buffer)
      buffer.SetValue(CByte(checkSum), 20)
      buffer.SetValue(CByte(&HF7), 21)

    Catch pEx As Exception

    End Try

    Return buffer

  End Function

#End Region

#Region "Zone Bass"

  ''' <summary>
  ''' Set Zone Turn On Bass
  ''' </summary>
  ''' <param name="cc"></param>
  ''' <param name="zz"></param>
  ''' <param name="volume"></param>
  Public Sub SetZoneBassEx(ByVal cc As Byte, ByVal zz As Byte, ByVal volume As Byte)
    AddCommand(SetZoneBass(cc, zz, volume))
    AddCommand(GetZoneBass(cc, zz))
  End Sub

  ''' <summary>
  ''' Set Zone Turn On Bass Increase
  ''' </summary>
  ''' <param name="cc"></param>
  ''' <param name="zz"></param>
  Public Sub SetZoneBassIncreaseEx(ByVal cc As Byte, ByVal zz As Byte)
    AddCommand(SetZoneBassIncrease(cc, zz))
    AddCommand(GetZoneBass(cc, zz))
  End Sub

  ''' <summary>
  ''' Set Zone Turn On Bass Decrease
  ''' </summary>
  ''' <param name="cc"></param>
  ''' <param name="zz"></param>
  Public Sub SetZoneBassDecreaseEx(ByVal cc As Byte, ByVal zz As Byte)
    AddCommand(SetZoneBassDecrease(cc, zz))
    AddCommand(GetZoneBass(cc, zz))
  End Sub

  ''' <summary>
  ''' Get Zone Volume
  ''' </summary>
  ''' <param name="cc"></param>
  ''' <param name="zz"></param>
  ''' <returns></returns>
  Protected Function GetZoneBass(ByVal cc As Byte, ByVal zz As Byte) As Byte()

    Dim buffer(17) As Byte

    Try
      Dim kk As Byte = IIf(m_DeviceModel = "ACA-E5", &H71, &H70)

      buffer.SetValue(CByte(&HF0), 0)       ' F0 Start of Message
      buffer.SetValue(CByte(cc), 1)         ' Target Controller Id
      buffer.SetValue(CByte(&H00), 2)       ' Target Zone Id
      buffer.SetValue(CByte(&H7F), 3)       ' Target Keypad Number

      buffer.SetValue(CByte(cc), 4)         ' Source Controller Id
      buffer.SetValue(CByte(zz), 5)         ' Source Zone Id
      buffer.SetValue(CByte(kk), 6)         ' Source Keypad Number

      buffer.SetValue(PacketMessageType.RequestData, 7) ' Message Type

      buffer.SetValue(CByte(&H05), 8)       ' Target Path:    0x04 (Represented by the following 5 bytes)
      buffer.SetValue(CByte(&H02), 9)       ' Path Level 0:   0x02 (Root Menu)
      buffer.SetValue(CByte(&H00), 10)      ' Path Level 1:   0x00 (Run Menu)
      buffer.SetValue(CByte(zz), 11)        ' Path Level 2:   zz   (Zone number -1)
      buffer.SetValue(CByte(&H00), 12)      ' Path Level 3:   0x00 (User Menu)
      buffer.SetValue(CByte(&H00), 13)      ' Path Level 4:   0x00 (Bass)
      buffer.SetValue(CByte(&H00), 14)      ' Source Path:    0x00 (this packet does not contain a source path)

      buffer.SetValue(CByte(&H00), 15)      ' 0x00

      Dim checkSum As Byte = CalculateChecksum(buffer)
      buffer.SetValue(CByte(checkSum), 16)
      buffer.SetValue(CByte(&HF7), 17)

    Catch pEx As Exception

    End Try

    Return buffer

  End Function
  ''' <summary>
  ''' Set Zone Bass Increase
  ''' </summary>
  ''' <param name="cc"></param>
  ''' <param name="zz"></param>
  ''' <returns></returns>
  Protected Function SetZoneBassIncrease(ByVal cc As Byte, ByVal zz As Byte) As Byte()

    Dim buffer(23) As Byte

    Try
      Dim kk As Byte = IIf(m_DeviceModel = "ACA-E5", &H71, &H70)

      buffer.SetValue(CByte(&HF0), 0)       ' F0 Start of Message
      buffer.SetValue(CByte(cc), 1)         ' Target Controller Id
      buffer.SetValue(CByte(&H00), 2)       ' Target Zone Id
      buffer.SetValue(CByte(&H7F), 3)       ' Target Keypad Number

      buffer.SetValue(CByte(cc), 4)         ' Source Controller Id
      buffer.SetValue(CByte(zz), 5)         ' Source Zone Id
      buffer.SetValue(CByte(kk), 6)         ' Source Keypad Number

      buffer.SetValue(PacketMessageType.Event, 7) ' Message Type (5 = Event)

      buffer.SetValue(CByte(&H05), 8)       ' 0x05
      buffer.SetValue(CByte(&H02), 9)       ' 0x02
      buffer.SetValue(CByte(&H00), 10)      ' 0x00
      buffer.SetValue(CByte(zz), 11)        ' Zone Number -1
      buffer.SetValue(CByte(&H00), 12)      ' 0x00
      buffer.SetValue(ZoneControlType.Bass, 13) ' 0x00
      buffer.SetValue(CByte(&H00), 14)      ' 0x00
      buffer.SetValue(CByte(&H69), 15)      ' 0x69 Increase Bass
      buffer.SetValue(CByte(&H00), 16)      ' 0x00
      buffer.SetValue(CByte(&H00), 17)      ' 0x00
      buffer.SetValue(CByte(&H00), 18)      ' 0x00
      buffer.SetValue(CByte(&H01), 19)      ' 0x01
      buffer.SetValue(CByte(&H00), 20)      ' 0x00
      buffer.SetValue(CByte(&H01), 21)      ' 0x01

      Dim checkSum As Byte = CalculateChecksum(buffer)
      buffer.SetValue(CByte(checkSum), 22)
      buffer.SetValue(CByte(&HF7), 23)

    Catch pEx As Exception

    End Try

    Return buffer

  End Function

  ''' <summary>
  ''' Zone Bass Decrease
  ''' </summary>
  ''' <param name="cc"></param>
  ''' <param name="zz"></param>
  ''' <returns></returns>
  Protected Function SetZoneBassDecrease(ByVal cc As Byte, ByVal zz As Byte) As Byte()

    Dim buffer(23) As Byte

    Try
      Dim kk As Byte = IIf(m_DeviceModel = "ACA-E5", &H71, &H70)

      buffer.SetValue(CByte(&HF0), 0)       ' F0 Start of Message
      buffer.SetValue(CByte(cc), 1)         ' Target Controller Id
      buffer.SetValue(CByte(&H00), 2)       ' Target Zone Id
      buffer.SetValue(CByte(&H7F), 3)       ' Target Keypad Number

      buffer.SetValue(CByte(cc), 4)         ' Source Controller Id
      buffer.SetValue(CByte(zz), 5)         ' Source Zone Id
      buffer.SetValue(CByte(kk), 6)         ' Source Keypad Number

      buffer.SetValue(PacketMessageType.Event, 7) ' Message Type (5 = Event)

      buffer.SetValue(CByte(&H05), 8)       ' 0x05
      buffer.SetValue(CByte(&H02), 9)       ' 0x02
      buffer.SetValue(CByte(&H00), 10)      ' 0x00
      buffer.SetValue(CByte(zz), 11)        ' Zone Number -1
      buffer.SetValue(CByte(&H00), 12)      ' 0x00
      buffer.SetValue(ZoneControlType.Bass, 13) ' 0x00
      buffer.SetValue(CByte(&H00), 14)      ' 0x00
      buffer.SetValue(CByte(&H6A), 15)      ' 0x69 Increase Decrease
      buffer.SetValue(CByte(&H00), 16)      ' 0x00
      buffer.SetValue(CByte(&H00), 17)      ' 0x00
      buffer.SetValue(CByte(&H00), 18)      ' 0x00
      buffer.SetValue(CByte(&H01), 19)      ' 0x01
      buffer.SetValue(CByte(&H00), 20)      ' 0x00
      buffer.SetValue(CByte(&H01), 21)      ' 0x01

      Dim checkSum As Byte = CalculateChecksum(buffer)
      buffer.SetValue(CByte(checkSum), 22)
      buffer.SetValue(CByte(&HF7), 23)

    Catch pEx As Exception

    End Try

    Return buffer

  End Function

  ''' <summary>
  ''' Set Zone Bass
  ''' </summary>
  ''' <param name="cc"></param>
  ''' <param name="zz"></param>
  ''' <param name="bass"></param>
  ''' <returns></returns>
  Protected Function SetZoneBass(ByVal cc As Byte, ByVal zz As Byte, ByVal bass As Byte) As Byte()

    Dim buffer(23) As Byte

    Try
      Dim kk As Byte = IIf(m_DeviceModel = "ACA-E5", &H71, &H70)

      buffer.SetValue(CByte(&HF0), 0)       ' F0 Start of Message
      buffer.SetValue(CByte(cc), 1)         ' Target Controller Id
      buffer.SetValue(CByte(&H00), 2)       ' Target Zone Id
      buffer.SetValue(CByte(&H7F), 3)       ' Target Keypad Number

      buffer.SetValue(CByte(cc), 4)         ' Source Controller Id
      buffer.SetValue(CByte(zz), 5)         ' Source Zone Id
      buffer.SetValue(CByte(kk), 6)         ' Source Keypad Number

      buffer.SetValue(PacketMessageType.SetData, 7) ' Message Type (0 = Set Data)

      buffer.SetValue(CByte(&H05), 8)       ' Target Path:  0x05 (Represented by the following 5 bytes)
      buffer.SetValue(CByte(&H02), 9)       ' Path Level 0: 0x00 (Root Menu)
      buffer.SetValue(CByte(&H00), 10)      ' Path Level 1: 0x00 (Run Menu)
      buffer.SetValue(CByte(zz), 11)        ' Path Level 2:   zz (Zone number -1)
      buffer.SetValue(CByte(&H00), 12)      ' Path Level 3: 0x00 (User Menu)
      buffer.SetValue(ZoneControlType.Bass, 13)  ' Path Level 4: 0x00 (Bass)
      buffer.SetValue(CByte(&H00), 14)      ' Source Path:  0x00 (This packet does not contain a source path)
      buffer.SetValue(CByte(&H00), 15)      ' 0x00
      buffer.SetValue(CByte(&H01), 16)      ' 0x00
      buffer.SetValue(CByte(&H01), 17)      ' 0x01
      buffer.SetValue(CByte(&H00), 18)      ' 0x00
      buffer.SetValue(CByte(&H01), 19)      ' 0x01
      buffer.SetValue(CByte(&H00), 20)      ' 0x00
      buffer.SetValue(CByte(bass), 21)      ' treble (0x00 = -10, 0x0A = Flat, 0x14 = +10)

      Dim checkSum As Byte = CalculateChecksum(buffer)
      buffer.SetValue(CByte(checkSum), 22)
      buffer.SetValue(CByte(&HF7), 23)

    Catch pEx As Exception

    End Try

    Return buffer

  End Function

#End Region

#Region "Zone Treble"

  ''' <summary>
  ''' Set Zone Turn On Treble
  ''' </summary>
  ''' <param name="cc"></param>
  ''' <param name="zz"></param>
  ''' <param name="volume"></param>
  Public Sub SetZoneTrebleEx(ByVal cc As Byte, ByVal zz As Byte, ByVal volume As Byte)
    AddCommand(SetZoneTreble(cc, zz, volume))
    AddCommand(GetZoneTreble(cc, zz))
  End Sub

  ''' <summary>
  ''' Set Zone Turn On Treble Increase
  ''' </summary>
  ''' <param name="cc"></param>
  ''' <param name="zz"></param>
  Public Sub SetZoneTrebleIncreaseEx(ByVal cc As Byte, ByVal zz As Byte)
    AddCommand(SetZoneTrebleIncrease(cc, zz))
    AddCommand(GetZoneTreble(cc, zz))
  End Sub

  ''' <summary>
  ''' Set Zone Turn On Treble Decrease
  ''' </summary>
  ''' <param name="cc"></param>
  ''' <param name="zz"></param>
  Public Sub SetZoneTrebleDecreaseEx(ByVal cc As Byte, ByVal zz As Byte)
    AddCommand(SetZoneTrebleDecrease(cc, zz))
    AddCommand(GetZoneTreble(cc, zz))
  End Sub

  ''' <summary>
  ''' Get Zone Volume
  ''' </summary>
  ''' <param name="cc"></param>
  ''' <param name="zz"></param>
  ''' <returns></returns>
  Protected Function GetZoneTreble(ByVal cc As Byte, ByVal zz As Byte) As Byte()

    Dim buffer(17) As Byte

    Try
      Dim kk As Byte = IIf(m_DeviceModel = "ACA-E5", &H71, &H70)

      buffer.SetValue(CByte(&HF0), 0)       ' F0 Start of Message
      buffer.SetValue(CByte(cc), 1)         ' Target Controller Id
      buffer.SetValue(CByte(&H00), 2)       ' Target Zone Id
      buffer.SetValue(CByte(&H7F), 3)       ' Target Keypad Number

      buffer.SetValue(CByte(cc), 4)         ' Source Controller Id
      buffer.SetValue(CByte(zz), 5)         ' Source Zone Id
      buffer.SetValue(CByte(kk), 6)         ' Source Keypad Number

      buffer.SetValue(PacketMessageType.RequestData, 7) ' Message Type

      buffer.SetValue(CByte(&H05), 8)       ' Target Path:    0x04 (Represented by the following 5 bytes)
      buffer.SetValue(CByte(&H02), 9)       ' Path Level 0:   0x02 (Root Menu)
      buffer.SetValue(CByte(&H00), 10)      ' Path Level 1:   0x00 (Run Menu)
      buffer.SetValue(CByte(zz), 11)        ' Path Level 2:   zz   (Zone number -1)
      buffer.SetValue(CByte(&H00), 12)      ' Path Level 3:   0x00 (User Menu)
      buffer.SetValue(CByte(&H01), 13)      ' Path Level 4:   0x01 (Treble)
      buffer.SetValue(CByte(&H00), 14)      ' Source Path:    0x00 (this packet does not contain a source path)

      buffer.SetValue(CByte(&H00), 15)      ' 0x00

      Dim checkSum As Byte = CalculateChecksum(buffer)
      buffer.SetValue(CByte(checkSum), 16)
      buffer.SetValue(CByte(&HF7), 17)

    Catch pEx As Exception

    End Try

    Return buffer

  End Function

  ''' <summary>
  ''' Set Zone Treble Increase
  ''' </summary>
  ''' <param name="cc"></param>
  ''' <param name="zz"></param>
  ''' <returns></returns>
  Protected Function SetZoneTrebleIncrease(ByVal cc As Byte, ByVal zz As Byte) As Byte()

    Dim buffer(23) As Byte

    Try
      Dim kk As Byte = IIf(m_DeviceModel = "ACA-E5", &H71, &H70)

      buffer.SetValue(CByte(&HF0), 0)       ' F0 Start of Message
      buffer.SetValue(CByte(cc), 1)         ' Target Controller Id
      buffer.SetValue(CByte(&H00), 2)       ' Target Zone Id
      buffer.SetValue(CByte(&H7F), 3)       ' Target Keypad Number

      buffer.SetValue(CByte(cc), 4)         ' Source Controller Id
      buffer.SetValue(CByte(zz), 5)         ' Source Zone Id
      buffer.SetValue(CByte(kk), 6)         ' Source Keypad Number

      buffer.SetValue(PacketMessageType.Event, 7) ' Message Type (5 = Event)

      buffer.SetValue(CByte(&H05), 8)       ' 0x05
      buffer.SetValue(CByte(&H02), 9)       ' 0x02
      buffer.SetValue(CByte(&H00), 10)      ' 0x00
      buffer.SetValue(CByte(zz), 11)        ' Zone Number -1
      buffer.SetValue(CByte(&H00), 12)      ' 0x00
      buffer.SetValue(ZoneControlType.Treble, 13) ' 0x01
      buffer.SetValue(CByte(&H00), 14)      ' 0x00
      buffer.SetValue(CByte(&H69), 15)      ' 0x69 Increase Treble
      buffer.SetValue(CByte(&H00), 16)      ' 0x00
      buffer.SetValue(CByte(&H00), 17)      ' 0x00
      buffer.SetValue(CByte(&H00), 18)      ' 0x00
      buffer.SetValue(CByte(&H02), 19)      ' 0x02
      buffer.SetValue(CByte(&H00), 20)      ' 0x00
      buffer.SetValue(CByte(&H01), 21)      ' 0x01

      Dim checkSum As Byte = CalculateChecksum(buffer)
      buffer.SetValue(CByte(checkSum), 22)
      buffer.SetValue(CByte(&HF7), 23)

    Catch pEx As Exception

    End Try

    Return buffer

  End Function

  ''' <summary>
  ''' Zone Treble Decrease
  ''' </summary>
  ''' <param name="cc"></param>
  ''' <param name="zz"></param>
  ''' <returns></returns>
  Protected Function SetZoneTrebleDecrease(ByVal cc As Byte, ByVal zz As Byte) As Byte()

    Dim buffer(23) As Byte

    Try
      Dim kk As Byte = IIf(m_DeviceModel = "ACA-E5", &H71, &H70)

      buffer.SetValue(CByte(&HF0), 0)       ' F0 Start of Message
      buffer.SetValue(CByte(cc), 1)         ' Target Controller Id
      buffer.SetValue(CByte(&H00), 2)       ' Target Zone Id
      buffer.SetValue(CByte(&H7F), 3)       ' Target Keypad Number

      buffer.SetValue(CByte(cc), 4)         ' Source Controller Id
      buffer.SetValue(CByte(zz), 5)         ' Source Zone Id
      buffer.SetValue(CByte(kk), 6)         ' Source Keypad Number

      buffer.SetValue(PacketMessageType.Event, 7) ' Message Type (5 = Event)

      buffer.SetValue(CByte(&H05), 8)       ' 0x05
      buffer.SetValue(CByte(&H02), 9)       ' 0x02
      buffer.SetValue(CByte(&H00), 10)      ' 0x00
      buffer.SetValue(CByte(zz), 11)        ' Zone Number -1
      buffer.SetValue(CByte(&H00), 12)      ' 0x00
      buffer.SetValue(ZoneControlType.Treble, 13)  ' 0x01
      buffer.SetValue(CByte(&H00), 14)      ' 0x00
      buffer.SetValue(CByte(&H6A), 15)      ' 0x69 Decrease Treble
      buffer.SetValue(CByte(&H00), 16)      ' 0x00
      buffer.SetValue(CByte(&H00), 17)      ' 0x00
      buffer.SetValue(CByte(&H00), 18)      ' 0x00
      buffer.SetValue(CByte(&H02), 19)      ' 0x02
      buffer.SetValue(CByte(&H00), 20)      ' 0x00
      buffer.SetValue(CByte(&H01), 21)      ' 0x01

      Dim checkSum As Byte = CalculateChecksum(buffer)
      buffer.SetValue(CByte(checkSum), 22)
      buffer.SetValue(CByte(&HF7), 23)

    Catch pEx As Exception

    End Try

    Return buffer

  End Function

  ''' <summary>
  ''' Set Zone Bass
  ''' </summary>
  ''' <param name="cc"></param>
  ''' <param name="zz"></param>
  ''' <param name="treble"></param>
  ''' <returns></returns>
  Protected Function SetZoneTreble(ByVal cc As Byte, ByVal zz As Byte, ByVal treble As Byte) As Byte()

    Dim buffer(23) As Byte

    Try
      Dim kk As Byte = IIf(m_DeviceModel = "ACA-E5", &H71, &H70)

      buffer.SetValue(CByte(&HF0), 0)       ' F0 Start of Message
      buffer.SetValue(CByte(cc), 1)         ' Target Controller Id
      buffer.SetValue(CByte(&H00), 2)       ' Target Zone Id
      buffer.SetValue(CByte(&H7F), 3)       ' Target Keypad Number

      buffer.SetValue(CByte(cc), 4)         ' Source Controller Id
      buffer.SetValue(CByte(zz), 5)         ' Source Zone Id
      buffer.SetValue(CByte(kk), 6)         ' Source Keypad Number

      buffer.SetValue(PacketMessageType.SetData, 7) ' Message Type (0 = Set Data)

      buffer.SetValue(CByte(&H05), 8)       ' Target Path:  0x05 (Represented by the following 5 bytes)
      buffer.SetValue(CByte(&H02), 9)       ' Path Level 0: 0x00 (Root Menu)
      buffer.SetValue(CByte(&H00), 10)      ' Path Level 1: 0x00 (Run Menu)
      buffer.SetValue(CByte(zz), 11)        ' Path Level 2:   zz (Zone number -1)
      buffer.SetValue(CByte(&H00), 12)      ' Path Level 3: 0x00 (User Menu)
      buffer.SetValue(ZoneControlType.Treble, 13)  ' Path Level 4: 0x01 (Treble)
      buffer.SetValue(CByte(&H00), 14)      ' Source Path:  0x00 (This packet does not contain a source path)
      buffer.SetValue(CByte(&H00), 15)      ' 0x00
      buffer.SetValue(CByte(&H01), 16)      ' 0x00
      buffer.SetValue(CByte(&H01), 17)      ' 0x01
      buffer.SetValue(CByte(&H00), 18)      ' 0x00
      buffer.SetValue(CByte(&H01), 19)      ' 0x01
      buffer.SetValue(CByte(&H00), 20)      ' 0x00
      buffer.SetValue(CByte(treble), 21)    ' treble (0x00 = -10, 0x0A = Flat, 0x14 = +10)

      Dim checkSum As Byte = CalculateChecksum(buffer)
      buffer.SetValue(CByte(checkSum), 22)
      buffer.SetValue(CByte(&HF7), 23)

    Catch pEx As Exception

    End Try

    Return buffer

  End Function

#End Region

#Region "Zone Loudness"

  ''' <summary>
  ''' Set Zone Turn On Loudness
  ''' </summary>
  ''' <param name="cc"></param>
  ''' <param name="zz"></param>
  ''' <param name="volume"></param>
  Public Sub SetZoneLoudnessEx(ByVal cc As Byte, ByVal zz As Byte, ByVal volume As Byte)
    AddCommand(SetZoneLoudness(cc, zz, volume))
    AddCommand(GetZoneLoudness(cc, zz))
  End Sub

  ''' <summary>
  ''' Set Zone Turn On Loudness Decrease
  ''' </summary>
  ''' <param name="cc"></param>
  ''' <param name="zz"></param>
  Public Sub SetZoneLoudnessToggleEx(ByVal cc As Byte, ByVal zz As Byte)
    AddCommand(SetZoneLoudnessToggle(cc, zz))
    AddCommand(GetZoneLoudness(cc, zz))
  End Sub

  ''' <summary>
  ''' Get Zone Loudness
  ''' </summary>
  ''' <param name="cc"></param>
  ''' <param name="zz"></param>
  ''' <returns></returns>
  Protected Function GetZoneLoudness(ByVal cc As Byte, ByVal zz As Byte) As Byte()

    Dim buffer(17) As Byte

    Try
      Dim kk As Byte = IIf(m_DeviceModel = "ACA-E5", &H71, &H70)

      buffer.SetValue(CByte(&HF0), 0)       ' F0 Start of Message
      buffer.SetValue(CByte(cc), 1)         ' Target Controller Id
      buffer.SetValue(CByte(&H00), 2)       ' Target Zone Id
      buffer.SetValue(CByte(&H7F), 3)       ' Target Keypad Number

      buffer.SetValue(CByte(cc), 4)         ' Source Controller Id
      buffer.SetValue(CByte(zz), 5)         ' Source Zone Id
      buffer.SetValue(CByte(kk), 6)         ' Source Keypad Number

      buffer.SetValue(PacketMessageType.RequestData, 7) ' Message Type

      buffer.SetValue(CByte(&H05), 8)       ' Target Path:    0x04 (Represented by the following 5 bytes)
      buffer.SetValue(CByte(&H02), 9)       ' Path Level 0:   0x02 (Root Menu)
      buffer.SetValue(CByte(&H00), 10)      ' Path Level 1:   0x00 (Run Menu)
      buffer.SetValue(CByte(zz), 11)        ' Path Level 2:   zz   (Zone number -1)
      buffer.SetValue(CByte(&H00), 12)      ' Path Level 3:   0x00 (User Menu)
      buffer.SetValue(CByte(&H02), 13)      ' Path Level 4:   0x02 (Loudness)
      buffer.SetValue(CByte(&H00), 14)      ' Source Path:    0x00 (this packet does not contain a source path)
      buffer.SetValue(CByte(&H00), 15)      ' 0x00

      Dim checkSum As Byte = CalculateChecksum(buffer)
      buffer.SetValue(CByte(checkSum), 16)
      buffer.SetValue(CByte(&HF7), 17)

    Catch pEx As Exception

    End Try

    Return buffer

  End Function

  ''' <summary>
  ''' Set Zone Loudness Toggle
  ''' </summary>
  ''' <param name="cc"></param>
  ''' <param name="zz"></param>
  ''' <returns></returns>
  Protected Function SetZoneLoudnessToggle(ByVal cc As Byte, ByVal zz As Byte) As Byte()

    Dim buffer(23) As Byte

    Try
      Dim kk As Byte = IIf(m_DeviceModel = "ACA-E5", &H71, &H70)

      buffer.SetValue(CByte(&HF0), 0)       ' F0 Start of Message
      buffer.SetValue(CByte(cc), 1)         ' Target Controller Id
      buffer.SetValue(CByte(&H00), 2)       ' Target Zone Id
      buffer.SetValue(CByte(&H7F), 3)       ' Target Keypad Number

      buffer.SetValue(CByte(cc), 4)         ' Source Controller Id
      buffer.SetValue(CByte(zz), 5)         ' Source Zone Id
      buffer.SetValue(CByte(kk), 6)         ' Source Keypad Number

      buffer.SetValue(PacketMessageType.Event, 7) ' Message Type (5 = Event)

      buffer.SetValue(CByte(&H05), 8)       ' 0x05
      buffer.SetValue(CByte(&H02), 9)       ' 0x02
      buffer.SetValue(CByte(&H00), 10)      ' 0x00
      buffer.SetValue(CByte(zz), 11)        ' Zone Number -1
      buffer.SetValue(CByte(&H00), 12)      ' 0x00
      buffer.SetValue(ZoneControlType.Loudness, 13) ' 0x02
      buffer.SetValue(CByte(&H00), 14)      ' 0x00
      buffer.SetValue(CByte(&H69), 15)      ' 0x69 On/Off Toggle
      buffer.SetValue(CByte(&H00), 16)      ' 0x00
      buffer.SetValue(CByte(&H00), 17)      ' 0x00
      buffer.SetValue(CByte(&H00), 18)      ' 0x00
      buffer.SetValue(CByte(&H03), 19)      ' 0x03
      buffer.SetValue(CByte(&H00), 20)      ' 0x00
      buffer.SetValue(CByte(&H01), 21)      ' 0x01

      Dim checkSum As Byte = CalculateChecksum(buffer)
      buffer.SetValue(CByte(checkSum), 22)
      buffer.SetValue(CByte(&HF7), 23)

    Catch pEx As Exception

    End Try

    Return buffer

  End Function

  ''' <summary>
  ''' Set Zone Loudness
  ''' </summary>
  ''' <param name="cc"></param>
  ''' <param name="zz"></param>
  ''' <param name="loudness"></param>
  ''' <returns></returns>
  Protected Function SetZoneLoudness(ByVal cc As Byte, ByVal zz As Byte, ByVal loudness As Byte) As Byte()

    Dim buffer(23) As Byte

    Try
      Dim kk As Byte = IIf(m_DeviceModel = "ACA-E5", &H71, &H70)

      buffer.SetValue(CByte(&HF0), 0)       ' F0 Start of Message
      buffer.SetValue(CByte(cc), 1)         ' Target Controller Id
      buffer.SetValue(CByte(&H00), 2)       ' Target Zone Id
      buffer.SetValue(CByte(&H7F), 3)       ' Target Keypad Number

      buffer.SetValue(CByte(cc), 4)         ' Source Controller Id
      buffer.SetValue(CByte(zz), 5)         ' Source Zone Id
      buffer.SetValue(CByte(kk), 6)         ' Source Keypad Number

      buffer.SetValue(PacketMessageType.SetData, 7) ' Message Type (0 = Set Data)

      buffer.SetValue(CByte(&H05), 8)       ' Target Path:  0x05 (Represented by the following 5 bytes)
      buffer.SetValue(CByte(&H02), 9)       ' Path Level 0: 0x00 (Root Menu)
      buffer.SetValue(CByte(&H00), 10)      ' Path Level 1: 0x00 (Run Menu)
      buffer.SetValue(CByte(zz), 11)        ' Path Level 2:   zz (Zone number -1)
      buffer.SetValue(CByte(&H00), 12)      ' Path Level 3: 0x00 (User Menu)
      buffer.SetValue(ZoneControlType.Loudness, 13)  ' Path Level 4: 0x02 (Loudness)
      buffer.SetValue(CByte(&H00), 14)      ' Source Path:  0x00 (This packet does not contain a source path)
      buffer.SetValue(CByte(&H00), 15)      ' 0x00
      buffer.SetValue(CByte(&H01), 16)      ' 0x00
      buffer.SetValue(CByte(&H01), 17)      ' 0x01
      buffer.SetValue(CByte(&H00), 18)      ' 0x00
      buffer.SetValue(CByte(&H01), 19)      ' 0x01
      buffer.SetValue(CByte(&H00), 20)      ' 0x00
      buffer.SetValue(CByte(loudness), 21)  ' loudness (0x00 = Off, 0x01 = On)

      Dim checkSum As Byte = CalculateChecksum(buffer)
      buffer.SetValue(CByte(checkSum), 22)
      buffer.SetValue(CByte(&HF7), 23)

    Catch pEx As Exception

    End Try

    Return buffer

  End Function

#End Region

#Region "Zone Balance"

  ''' <summary>
  ''' Set Zone Turn On Balance
  ''' </summary>
  ''' <param name="cc"></param>
  ''' <param name="zz"></param>
  ''' <param name="volume"></param>
  Public Sub SetZoneBalanceEx(ByVal cc As Byte, ByVal zz As Byte, ByVal volume As Byte)
    AddCommand(SetZoneBalance(cc, zz, volume))
    AddCommand(GetZoneBalance(cc, zz))
  End Sub

  ''' <summary>
  ''' Set Zone Turn On Balance Left
  ''' </summary>
  ''' <param name="cc"></param>
  ''' <param name="zz"></param>
  Public Sub SetZoneBalanceLeftEx(ByVal cc As Byte, ByVal zz As Byte)
    AddCommand(SetZoneBalanceLeft(cc, zz))
    AddCommand(GetZoneBalance(cc, zz))
  End Sub

  ''' <summary>
  ''' Set Zone Turn On Balance Right
  ''' </summary>
  ''' <param name="cc"></param>
  ''' <param name="zz"></param>
  Public Sub SetZoneBalanceRightEx(ByVal cc As Byte, ByVal zz As Byte)
    AddCommand(SetZoneBalanceRight(cc, zz))
    AddCommand(GetZoneBalance(cc, zz))
  End Sub

  ''' <summary>
  ''' Get Zone Balance
  ''' </summary>
  ''' <param name="cc"></param>
  ''' <param name="zz"></param>
  ''' <returns></returns>
  Protected Function GetZoneBalance(ByVal cc As Byte, ByVal zz As Byte) As Byte()

    Dim buffer(17) As Byte

    Try
      Dim kk As Byte = IIf(m_DeviceModel = "ACA-E5", &H71, &H70)

      buffer.SetValue(CByte(&HF0), 0)       ' F0 Start of Message
      buffer.SetValue(CByte(cc), 1)         ' Target Controller Id
      buffer.SetValue(CByte(&H00), 2)       ' Target Zone Id
      buffer.SetValue(CByte(&H7F), 3)       ' Target Keypad Number

      buffer.SetValue(CByte(cc), 4)         ' Source Controller Id
      buffer.SetValue(CByte(zz), 5)         ' Source Zone Id
      buffer.SetValue(CByte(kk), 6)         ' Source Keypad Number

      buffer.SetValue(PacketMessageType.RequestData, 7) ' Message Type

      buffer.SetValue(CByte(&H05), 8)       ' Target Path:    0x05 (Represented by the following 5 bytes)
      buffer.SetValue(CByte(&H02), 9)       ' Path Level 0:   0x02 (Root Menu)
      buffer.SetValue(CByte(&H00), 10)      ' Path Level 1:   0x00 (Run Menu)
      buffer.SetValue(CByte(zz), 11)        ' Path Level 2:   zz   (Zone number -1)
      buffer.SetValue(CByte(&H00), 12)      ' Path Level 3:   0x00 (User Menu)
      buffer.SetValue(CByte(&H03), 13)      ' Path Level 4:   0x03 (Balance)
      buffer.SetValue(CByte(&H00), 14)      ' Source Path:    0x00 (this packet does not contain a source path)
      buffer.SetValue(CByte(&H00), 15)      ' 0x00

      Dim checkSum As Byte = CalculateChecksum(buffer)
      buffer.SetValue(CByte(checkSum), 16)
      buffer.SetValue(CByte(&HF7), 17)

    Catch pEx As Exception

    End Try

    Return buffer

  End Function

  ''' <summary>
  ''' Set Zone Balance Left
  ''' </summary>
  ''' <param name="cc"></param>
  ''' <param name="zz"></param>
  ''' <returns></returns>
  Protected Function SetZoneBalanceLeft(ByVal cc As Byte, ByVal zz As Byte) As Byte()

    Dim buffer(23) As Byte

    Try
      Dim kk As Byte = IIf(m_DeviceModel = "ACA-E5", &H71, &H70)

      buffer.SetValue(CByte(&HF0), 0)       ' F0 Start of Message
      buffer.SetValue(CByte(cc), 1)         ' Target Controller Id
      buffer.SetValue(CByte(&H00), 2)       ' Target Zone Id
      buffer.SetValue(CByte(&H7F), 3)       ' Target Keypad Number

      buffer.SetValue(CByte(cc), 4)         ' Source Controller Id
      buffer.SetValue(CByte(zz), 5)         ' Source Zone Id
      buffer.SetValue(CByte(kk), 6)         ' Source Keypad Number

      buffer.SetValue(PacketMessageType.Event, 7) ' Message Type (5 = Event)

      buffer.SetValue(CByte(&H05), 8)       ' 0x05
      buffer.SetValue(CByte(&H02), 9)       ' 0x02
      buffer.SetValue(CByte(&H00), 10)      ' 0x00
      buffer.SetValue(CByte(zz), 11)        ' Zone Number -1
      buffer.SetValue(CByte(&H00), 12)      ' 0x00
      buffer.SetValue(ZoneControlType.Balance, 13) ' 0x03
      buffer.SetValue(CByte(&H00), 14)      ' 0x00
      buffer.SetValue(CByte(&H69), 15)      ' 0x69 More Left
      buffer.SetValue(CByte(&H00), 16)      ' 0x00
      buffer.SetValue(CByte(&H00), 17)      ' 0x00
      buffer.SetValue(CByte(&H00), 18)      ' 0x00
      buffer.SetValue(CByte(&H04), 19)      ' 0x04
      buffer.SetValue(CByte(&H00), 20)      ' 0x00
      buffer.SetValue(CByte(&H01), 21)      ' 0x01

      Dim checkSum As Byte = CalculateChecksum(buffer)
      buffer.SetValue(CByte(checkSum), 22)
      buffer.SetValue(CByte(&HF7), 23)

    Catch pEx As Exception

    End Try

    Return buffer

  End Function

  ''' <summary>
  ''' Zone Treble Decrease
  ''' </summary>
  ''' <param name="cc"></param>
  ''' <param name="zz"></param>
  ''' <returns></returns>
  Protected Function SetZoneBalanceRight(ByVal cc As Byte, ByVal zz As Byte) As Byte()

    Dim buffer(23) As Byte

    Try
      Dim kk As Byte = IIf(m_DeviceModel = "ACA-E5", &H71, &H70)

      buffer.SetValue(CByte(&HF0), 0)       ' F0 Start of Message
      buffer.SetValue(CByte(cc), 1)         ' Target Controller Id
      buffer.SetValue(CByte(&H00), 2)       ' Target Zone Id
      buffer.SetValue(CByte(&H7F), 3)       ' Target Keypad Number

      buffer.SetValue(CByte(cc), 4)         ' Source Controller Id
      buffer.SetValue(CByte(zz), 5)         ' Source Zone Id
      buffer.SetValue(CByte(kk), 6)         ' Source Keypad Number

      buffer.SetValue(PacketMessageType.Event, 7) ' Message Type (5 = Event)

      buffer.SetValue(CByte(&H05), 8)       ' 0x05
      buffer.SetValue(CByte(&H02), 9)       ' 0x02
      buffer.SetValue(CByte(&H00), 10)      ' 0x00
      buffer.SetValue(CByte(zz), 11)        ' Zone Number -1
      buffer.SetValue(CByte(&H00), 12)      ' 0x00
      buffer.SetValue(ZoneControlType.Balance, 13) ' 0x03
      buffer.SetValue(CByte(&H00), 14)      ' 0x00
      buffer.SetValue(CByte(&H6A), 15)      ' 0x6A More Right
      buffer.SetValue(CByte(&H00), 16)      ' 0x00
      buffer.SetValue(CByte(&H00), 17)      ' 0x00
      buffer.SetValue(CByte(&H00), 18)      ' 0x00
      buffer.SetValue(CByte(&H04), 19)      ' 0x04
      buffer.SetValue(CByte(&H00), 20)      ' 0x00
      buffer.SetValue(CByte(&H01), 21)      ' 0x01

      Dim checkSum As Byte = CalculateChecksum(buffer)
      buffer.SetValue(CByte(checkSum), 22)
      buffer.SetValue(CByte(&HF7), 23)

    Catch pEx As Exception

    End Try

    Return buffer

  End Function

  ''' <summary>
  ''' Set Zone Bass
  ''' </summary>
  ''' <param name="cc"></param>
  ''' <param name="zz"></param>
  ''' <param name="balance"></param>
  ''' <returns></returns>
  Protected Function SetZoneBalance(ByVal cc As Byte, ByVal zz As Byte, ByVal balance As Byte) As Byte()

    Dim buffer(23) As Byte

    Try
      Dim kk As Byte = IIf(m_DeviceModel = "ACA-E5", &H71, &H70)

      buffer.SetValue(CByte(&HF0), 0)       ' F0 Start of Message
      buffer.SetValue(CByte(cc), 1)         ' Target Controller Id
      buffer.SetValue(CByte(&H00), 2)       ' Target Zone Id
      buffer.SetValue(CByte(&H7F), 3)       ' Target Keypad Number

      buffer.SetValue(CByte(cc), 4)         ' Source Controller Id
      buffer.SetValue(CByte(zz), 5)         ' Source Zone Id
      buffer.SetValue(CByte(kk), 6)         ' Source Keypad Number

      buffer.SetValue(PacketMessageType.SetData, 7) ' Message Type (0 = Set Data)

      buffer.SetValue(CByte(&H05), 8)       ' Target Path:  0x05 (Represented by the following 5 bytes)
      buffer.SetValue(CByte(&H02), 9)       ' Path Level 0: 0x00 (Root Menu)
      buffer.SetValue(CByte(&H00), 10)      ' Path Level 1: 0x00 (Run Menu)
      buffer.SetValue(CByte(zz), 11)        ' Path Level 2:   zz (Zone number -1)
      buffer.SetValue(CByte(&H00), 12)      ' Path Level 3: 0x00 (User Menu)
      buffer.SetValue(ZoneControlType.Balance, 13)  ' Path Level 4: 0x03 (Balance)
      buffer.SetValue(CByte(&H00), 14)      ' Source Path:  0x00 (This packet does not contain a source path)
      buffer.SetValue(CByte(&H00), 15)      ' 0x00
      buffer.SetValue(CByte(&H01), 16)      ' 0x00
      buffer.SetValue(CByte(&H01), 17)      ' 0x01
      buffer.SetValue(CByte(&H00), 18)      ' 0x00
      buffer.SetValue(CByte(&H01), 19)      ' 0x01
      buffer.SetValue(CByte(&H00), 20)      ' 0x00
      buffer.SetValue(CByte(balance), 21)   ' balance (0x00 = More Left, 0x0A = Middle, 0x14 = More Right)

      Dim checkSum As Byte = CalculateChecksum(buffer)
      buffer.SetValue(CByte(checkSum), 22)
      buffer.SetValue(CByte(&HF7), 23)

    Catch pEx As Exception

    End Try

    Return buffer

  End Function

#End Region

#Region "Zone Turn On Volume"

  ''' <summary>
  ''' Set Zone Turn On Volume
  ''' </summary>
  ''' <param name="cc"></param>
  ''' <param name="zz"></param>
  ''' <param name="volume"></param>
  Public Sub SetZoneTurnOnVolumeEx(ByVal cc As Byte, ByVal zz As Byte, ByVal volume As Byte)
    AddCommand(SetZoneTurnOnVolume(cc, zz, volume))
    AddCommand(GetZoneTurnOnVolume(cc, zz))
  End Sub

  ''' <summary>
  ''' Set Zone Turn On Volume Increase
  ''' </summary>
  ''' <param name="cc"></param>
  ''' <param name="zz"></param>
  Public Sub SetZoneTurnOnVolumeIncreaseEx(ByVal cc As Byte, ByVal zz As Byte)
    AddCommand(SetZoneTurnOnVolumeIncrease(cc, zz))
    AddCommand(GetZoneTurnOnVolume(cc, zz))
  End Sub

  ''' <summary>
  ''' Set Zone Turn On Volume Decrease
  ''' </summary>
  ''' <param name="cc"></param>
  ''' <param name="zz"></param>
  Public Sub SetZoneTurnOnVolumeDecreaseEx(ByVal cc As Byte, ByVal zz As Byte)
    AddCommand(SetZoneTurnOnVolumeDecrease(cc, zz))
    AddCommand(GetZoneTurnOnVolume(cc, zz))
  End Sub

  ''' <summary>
  ''' Get Zone Turn On Volume
  ''' </summary>
  ''' <param name="cc"></param>
  ''' <param name="zz"></param>
  ''' <returns></returns>
  Protected Function GetZoneTurnOnVolume(ByVal cc As Byte, ByVal zz As Byte) As Byte()

    Dim buffer(17) As Byte

    Try
      Dim kk As Byte = IIf(m_DeviceModel = "ACA-E5", &H71, &H70)

      buffer.SetValue(CByte(&HF0), 0)       ' F0 Start of Message
      buffer.SetValue(CByte(cc), 1)         ' Target Controller Id
      buffer.SetValue(CByte(&H00), 2)       ' Target Zone Id
      buffer.SetValue(CByte(&H7F), 3)       ' Target Keypad Number

      buffer.SetValue(CByte(cc), 4)         ' Source Controller Id
      buffer.SetValue(CByte(zz), 5)         ' Source Zone Id
      buffer.SetValue(CByte(kk), 6)         ' Source Keypad Number

      buffer.SetValue(PacketMessageType.RequestData, 7) ' Message Type

      buffer.SetValue(CByte(&H05), 8)       ' Target Path:    0x04 (Represented by the following 5 bytes)
      buffer.SetValue(CByte(&H02), 9)       ' Path Level 0:   0x02 (Root Menu)
      buffer.SetValue(CByte(&H00), 10)      ' Path Level 1:   0x00 (Run Menu)
      buffer.SetValue(CByte(zz), 11)        ' Path Level 2:   zz   (Zone number -1)
      buffer.SetValue(CByte(&H00), 12)      ' Path Level 3:   0x00 (User Menu)
      buffer.SetValue(CByte(&H04), 13)      ' Path Level 4:   0x04 (TURN ON VOL)
      buffer.SetValue(CByte(&H00), 14)      ' Source Path:    0x00 (this packet does not contain a source path)
      buffer.SetValue(CByte(&H00), 15)      ' 0x00

      Dim checkSum As Byte = CalculateChecksum(buffer)
      buffer.SetValue(CByte(checkSum), 16)
      buffer.SetValue(CByte(&HF7), 17)

    Catch pEx As Exception

    End Try

    Return buffer

  End Function

  ''' <summary>
  ''' Set Zone Turn On Volume Increase
  ''' </summary>
  ''' <param name="cc"></param>
  ''' <param name="zz"></param>
  ''' <returns></returns>
  Protected Function SetZoneTurnOnVolumeIncrease(ByVal cc As Byte, ByVal zz As Byte) As Byte()

    Dim buffer(23) As Byte

    Try
      Dim kk As Byte = IIf(m_DeviceModel = "ACA-E5", &H71, &H70)

      buffer.SetValue(CByte(&HF0), 0)       ' F0 Start of Message
      buffer.SetValue(CByte(cc), 1)         ' Target Controller Id
      buffer.SetValue(CByte(&H00), 2)       ' Target Zone Id
      buffer.SetValue(CByte(&H7F), 3)       ' Target Keypad Number

      buffer.SetValue(CByte(cc), 4)         ' Source Controller Id
      buffer.SetValue(CByte(zz), 5)         ' Source Zone Id
      buffer.SetValue(CByte(kk), 6)         ' Source Keypad Number

      buffer.SetValue(PacketMessageType.Event, 7) ' Message Type (5 = Event)

      buffer.SetValue(CByte(&H05), 8)       ' 0x05
      buffer.SetValue(CByte(&H02), 9)       ' 0x02
      buffer.SetValue(CByte(&H00), 10)      ' 0x00
      buffer.SetValue(CByte(zz), 11)        ' Zone Number -1
      buffer.SetValue(CByte(&H00), 12)      ' 0x00
      buffer.SetValue(ZoneControlType.TurnOnVolume, 13) ' 0x04
      buffer.SetValue(CByte(&H00), 14)      ' 0x00
      buffer.SetValue(CByte(&H69), 15)      ' 0x69 More Left
      buffer.SetValue(CByte(&H00), 16)      ' 0x00
      buffer.SetValue(CByte(&H00), 17)      ' 0x00
      buffer.SetValue(CByte(&H00), 18)      ' 0x00
      buffer.SetValue(CByte(&H05), 19)      ' 0x05
      buffer.SetValue(CByte(&H00), 20)      ' 0x00
      buffer.SetValue(CByte(&H01), 21)      ' 0x01

      Dim checkSum As Byte = CalculateChecksum(buffer)
      buffer.SetValue(CByte(checkSum), 22)
      buffer.SetValue(CByte(&HF7), 23)

    Catch pEx As Exception

    End Try

    Return buffer

  End Function

  ''' <summary>
  ''' Set Zone Turn On Volume Decrease
  ''' </summary>
  ''' <param name="cc"></param>
  ''' <param name="zz"></param>
  ''' <returns></returns>
  Protected Function SetZoneTurnOnVolumeDecrease(ByVal cc As Byte, ByVal zz As Byte) As Byte()

    Dim buffer(23) As Byte

    Try
      Dim kk As Byte = IIf(m_DeviceModel = "ACA-E5", &H71, &H70)

      buffer.SetValue(CByte(&HF0), 0)       ' F0 Start of Message
      buffer.SetValue(CByte(cc), 1)         ' Target Controller Id
      buffer.SetValue(CByte(&H00), 2)       ' Target Zone Id
      buffer.SetValue(CByte(&H7F), 3)       ' Target Keypad Number

      buffer.SetValue(CByte(cc), 4)         ' Source Controller Id
      buffer.SetValue(CByte(zz), 5)         ' Source Zone Id
      buffer.SetValue(CByte(kk), 6)         ' Source Keypad Number

      buffer.SetValue(PacketMessageType.Event, 7) ' Message Type (5 = Event)

      buffer.SetValue(CByte(&H05), 8)       ' 0x05
      buffer.SetValue(CByte(&H02), 9)       ' 0x02
      buffer.SetValue(CByte(&H00), 10)      ' 0x00
      buffer.SetValue(CByte(zz), 11)        ' Zone Number -1
      buffer.SetValue(CByte(&H00), 12)      ' 0x00
      buffer.SetValue(ZoneControlType.TurnOnVolume, 13) ' 0x04
      buffer.SetValue(CByte(&H00), 14)      ' 0x00
      buffer.SetValue(CByte(&H6A), 15)      ' 0x69 More Right
      buffer.SetValue(CByte(&H00), 16)      ' 0x00
      buffer.SetValue(CByte(&H00), 17)      ' 0x00
      buffer.SetValue(CByte(&H00), 18)      ' 0x00
      buffer.SetValue(CByte(&H05), 19)      ' 0x05
      buffer.SetValue(CByte(&H00), 20)      ' 0x00
      buffer.SetValue(CByte(&H01), 21)      ' 0x01

      Dim checkSum As Byte = CalculateChecksum(buffer)
      buffer.SetValue(CByte(checkSum), 22)
      buffer.SetValue(CByte(&HF7), 23)

    Catch pEx As Exception

    End Try

    Return buffer

  End Function

  ''' <summary>
  ''' Set Zone Turn On Volume
  ''' </summary>
  ''' <param name="cc"></param>
  ''' <param name="zz"></param>
  ''' <param name="volume"></param>
  ''' <returns></returns>
  Protected Function SetZoneTurnOnVolume(ByVal cc As Byte, ByVal zz As Byte, ByVal volume As Byte) As Byte()

    Dim buffer(23) As Byte

    Try
      Dim kk As Byte = IIf(m_DeviceModel = "ACA-E5", &H71, &H70)

      buffer.SetValue(CByte(&HF0), 0)       ' F0 Start of Message
      buffer.SetValue(CByte(cc), 1)         ' Target Controller Id
      buffer.SetValue(CByte(&H00), 2)       ' Target Zone Id
      buffer.SetValue(CByte(&H7F), 3)       ' Target Keypad Number

      buffer.SetValue(CByte(cc), 4)         ' Source Controller Id
      buffer.SetValue(CByte(zz), 5)         ' Source Zone Id
      buffer.SetValue(CByte(kk), 6)         ' Source Keypad Number

      buffer.SetValue(PacketMessageType.SetData, 7) ' Message Type (0 = Set Data)

      buffer.SetValue(CByte(&H05), 8)       ' Target Path:  0x05 (Represented by the following 5 bytes)
      buffer.SetValue(CByte(&H02), 9)       ' Path Level 0: 0x02 (Root Menu)
      buffer.SetValue(CByte(&H00), 10)      ' Path Level 1: 0x00 (Run Menu)
      buffer.SetValue(CByte(zz), 11)        ' Path Level 2:   zz (Zone number -1)
      buffer.SetValue(CByte(&H00), 12)      ' Path Level 3: 0x00 (User Menu)
      buffer.SetValue(ZoneControlType.TurnOnVolume, 13)  ' Path Level 4: 0x04 (TurnOnVolume)
      buffer.SetValue(CByte(&H00), 14)      ' Source Path:  0x00 (This packet does not contain a source path)
      buffer.SetValue(CByte(&H00), 15)      ' 0x00
      buffer.SetValue(CByte(&H01), 16)      ' 0x00
      buffer.SetValue(CByte(&H01), 17)      ' 0x01
      buffer.SetValue(CByte(&H00), 18)      ' 0x00
      buffer.SetValue(CByte(&H01), 19)      ' 0x01
      buffer.SetValue(CByte(&H00), 20)      ' 0x00
      buffer.SetValue(CByte(volume), 21)    ' volume (0x00 - 0x32, 0x00 = 0 ... 0x32 = 100)

      Dim checkSum As Byte = CalculateChecksum(buffer)
      buffer.SetValue(CByte(checkSum), 22)
      buffer.SetValue(CByte(&HF7), 23)

    Catch pEx As Exception

    End Try

    Return buffer

  End Function

#End Region

#Region "Zone Background Color"

  ''' <summary>
  ''' Set Zone Keypad Background Color
  ''' </summary>
  ''' <param name="cc"></param>
  ''' <param name="zz"></param>
  ''' <param name="color"></param>
  Public Sub SetZoneBackgroundColorEx(ByVal cc As Byte, ByVal zz As Byte, ByVal color As Byte)
    AddCommand(SetZoneBackgroundColor(cc, zz, color))
    AddCommand(GetZoneBackgroundColor(cc, zz))
  End Sub

  ''' <summary>
  ''' Set Zone Keypad Background Color Toggle
  ''' </summary>
  ''' <param name="cc"></param>
  ''' <param name="zz"></param>
  Public Sub SetZoneBackgroundColorToggleEx(ByVal cc As Byte, ByVal zz As Byte)
    AddCommand(SetZoneBackGroundColorToggle(cc, zz))
    AddCommand(GetZoneBackgroundColor(cc, zz))
  End Sub

  ''' <summary>
  ''' Set Zone Keypad Background Color
  ''' </summary>
  ''' <param name="cc"></param>
  ''' <param name="zz"></param>
  ''' <returns></returns>
  Protected Function SetZoneBackGroundColorToggle(ByVal cc As Byte, ByVal zz As Byte) As Byte()

    Dim buffer(23) As Byte

    Try
      Dim kk As Byte = IIf(m_DeviceModel = "ACA-E5", &H71, &H70)

      buffer.SetValue(CByte(&HF0), 0)       ' F0 Start of Message
      buffer.SetValue(CByte(cc), 1)         ' Target Controller Id
      buffer.SetValue(CByte(&H00), 2)       ' Target Zone Id
      buffer.SetValue(CByte(&H7F), 3)       ' Target Keypad Number

      buffer.SetValue(CByte(cc), 4)         ' Source Controller Id
      buffer.SetValue(CByte(zz), 5)         ' Source Zone Id
      buffer.SetValue(CByte(kk), 6)         ' Source Keypad Number

      buffer.SetValue(PacketMessageType.Event, 7) ' Message Type (5 = Event)

      buffer.SetValue(CByte(&H05), 8)       ' 0x05
      buffer.SetValue(CByte(&H02), 9)       ' 0x02
      buffer.SetValue(CByte(&H00), 10)      ' 0x00
      buffer.SetValue(CByte(zz), 11)        ' Zone Number -1
      buffer.SetValue(CByte(&H00), 12)      ' 0x00
      buffer.SetValue(ZoneControlType.BackgroundColor, 13) ' 0x05
      buffer.SetValue(CByte(&H00), 14)      ' 0x00
      buffer.SetValue(CByte(&H69), 15)      ' 0x69 Background Color Toggle
      buffer.SetValue(CByte(&H00), 16)      ' 0x00
      buffer.SetValue(CByte(&H00), 17)      ' 0x00
      buffer.SetValue(CByte(&H00), 18)      ' 0x00
      buffer.SetValue(CByte(&H05), 19)      ' 0x06
      buffer.SetValue(CByte(&H00), 20)      ' 0x00
      buffer.SetValue(CByte(&H01), 21)      ' 0x01

      Dim checkSum As Byte = CalculateChecksum(buffer)
      buffer.SetValue(CByte(checkSum), 22)
      buffer.SetValue(CByte(&HF7), 23)

    Catch pEx As Exception

    End Try

    Return buffer

  End Function

  ''' <summary>
  ''' Set Zone Background Color
  ''' </summary>
  ''' <param name="cc"></param>
  ''' <param name="zz"></param>
  ''' <param name="color"></param>
  ''' <returns></returns>
  Protected Function SetZoneBackgroundColor(ByVal cc As Byte, ByVal zz As Byte, ByVal color As Byte) As Byte()

    Dim buffer(23) As Byte

    Try
      Dim kk As Byte = IIf(m_DeviceModel = "ACA-E5", &H71, &H70)

      buffer.SetValue(CByte(&HF0), 0)       ' F0 Start of Message
      buffer.SetValue(CByte(cc), 1)         ' Target Controller Id
      buffer.SetValue(CByte(&H00), 2)       ' Target Zone Id
      buffer.SetValue(CByte(&H7F), 3)       ' Target Keypad Number

      buffer.SetValue(CByte(cc), 4)         ' Source Controller Id
      buffer.SetValue(CByte(zz), 5)         ' Source Zone Id
      buffer.SetValue(CByte(kk), 6)         ' Source Keypad Number

      buffer.SetValue(PacketMessageType.SetData, 7) ' Message Type (0 = Set Data)

      buffer.SetValue(CByte(&H05), 8)       ' 0x05
      buffer.SetValue(CByte(&H02), 9)       ' 0x02
      buffer.SetValue(CByte(&H00), 10)      ' 0x00
      buffer.SetValue(CByte(zz), 11)        ' Zone Number -1
      buffer.SetValue(CByte(&H00), 12)      ' 0x00
      buffer.SetValue(ZoneControlType.BackgroundColor, 13) ' 0x05
      buffer.SetValue(CByte(&H00), 14)      ' 0x00
      buffer.SetValue(CByte(&H00), 15)      ' 0x00
      buffer.SetValue(CByte(&H01), 16)      ' 0x00
      buffer.SetValue(CByte(&H01), 17)      ' 0x01
      buffer.SetValue(CByte(&H00), 18)      ' 0x00
      buffer.SetValue(CByte(&H01), 19)      ' 0x01
      buffer.SetValue(CByte(&H00), 20)      ' 0x00
      buffer.SetValue(CByte(color), 21)     ' color (0x00 = Off, 0x01 = Amber, 0x02 = Green)

      Dim checkSum As Byte = CalculateChecksum(buffer)
      buffer.SetValue(CByte(checkSum), 22)
      buffer.SetValue(CByte(&HF7), 23)

    Catch pEx As Exception

    End Try

    Return buffer

  End Function

  ''' <summary>
  ''' Get Zone Background Color
  ''' </summary>
  ''' <param name="cc"></param>
  ''' <param name="zz"></param>
  ''' <returns></returns>
  Protected Function GetZoneBackgroundColor(ByVal cc As Byte, ByVal zz As Byte) As Byte()

    Dim buffer(17) As Byte

    Try
      Dim kk As Byte = IIf(m_DeviceModel = "ACA-E5", &H71, &H70)

      buffer.SetValue(CByte(&HF0), 0)       ' F0 Start of Message
      buffer.SetValue(CByte(cc), 1)         ' Target Controller Id
      buffer.SetValue(CByte(&H00), 2)       ' Target Zone Id
      buffer.SetValue(CByte(&H7F), 3)       ' Target Keypad Number

      buffer.SetValue(CByte(cc), 4)         ' Source Controller Id
      buffer.SetValue(CByte(zz), 5)         ' Source Zone Id
      buffer.SetValue(CByte(kk), 6)         ' Source Keypad Number

      buffer.SetValue(PacketMessageType.RequestData, 7) ' Message Type

      buffer.SetValue(CByte(&H05), 8)       ' 0x05
      buffer.SetValue(CByte(&H02), 9)       ' 0x02
      buffer.SetValue(CByte(&H00), 10)      ' 0x00
      buffer.SetValue(CByte(zz), 11)        ' Zone number -1
      buffer.SetValue(CByte(&H00), 12)      ' 0x00
      buffer.SetValue(ZoneControlType.BackgroundColor, 13) ' 0x05
      buffer.SetValue(CByte(&H00), 14)      ' 0x00
      buffer.SetValue(CByte(&H00), 15)      ' 0x00

      Dim checkSum As Byte = CalculateChecksum(buffer)
      buffer.SetValue(CByte(checkSum), 16)
      buffer.SetValue(CByte(&HF7), 17)

    Catch pEx As Exception

    End Try

    Return buffer

  End Function

#End Region

#Region "Zone Do Not Disturb"

  ''' <summary>
  ''' Set Zone Do Not Disturb Ex
  ''' </summary>
  ''' <param name="cc"></param>
  ''' <param name="zz"></param>
  ''' <param name="dnd"></param>
  Public Sub SetZoneDoNotDisturbEx(ByVal cc As Byte, ByVal zz As Byte, ByVal dnd As Byte)
    AddCommand(SetZoneDoNotDisturb(cc, zz, dnd))
    AddCommand(GetZoneStatus(cc, zz))
  End Sub

  ''' <summary>
  ''' Set Zone Do Not Disturb Toggle Ex
  ''' </summary>
  ''' <param name="cc"></param>
  ''' <param name="zz"></param>
  Public Sub SetZoneDoNotDisturbToggleEx(ByVal cc As Byte, ByVal zz As Byte)
    AddCommand(SetZoneDoNotDisturbToggle(cc, zz))
  End Sub

  ''' <summary>
  ''' Get Zone Do Not Disturb
  ''' </summary>
  ''' <param name="cc"></param>
  ''' <param name="zz"></param>
  ''' <returns></returns>
  Protected Function GetZoneDoNotDisturb(ByVal cc As Byte, ByVal zz As Byte) As Byte()

    Dim buffer(17) As Byte

    Try
      Dim kk As Byte = IIf(m_DeviceModel = "ACA-E5", &H71, &H70)

      buffer.SetValue(CByte(&HF0), 0)       ' F0 Start of Message
      buffer.SetValue(CByte(cc), 1)         ' Target Controller Id
      buffer.SetValue(CByte(&H00), 2)       ' Target Zone Id
      buffer.SetValue(CByte(&H7F), 3)       ' Target Keypad Number

      buffer.SetValue(CByte(cc), 4)         ' Source Controller Id
      buffer.SetValue(CByte(zz), 5)         ' Source Zone Id
      buffer.SetValue(CByte(kk), 6)         ' Source Keypad Number

      buffer.SetValue(PacketMessageType.RequestData, 7) ' Message Type

      buffer.SetValue(CByte(&H05), 8)       ' Target Path:    0x04 (Represented by the following 5 bytes)
      buffer.SetValue(CByte(&H02), 9)       ' Path Level 0:   0x02 (Root Menu)
      buffer.SetValue(CByte(&H00), 10)      ' Path Level 1:   0x00 (Run Menu)
      buffer.SetValue(CByte(zz), 11)        ' Path Level 2:   zz   (Zone number -1)
      buffer.SetValue(CByte(&H00), 12)      ' Path Level 3:   0x00 (User Menu)
      buffer.SetValue(CByte(&H06), 13)      ' Path Level 4:   0x06 (DO NOT DSTRB)
      buffer.SetValue(CByte(&H00), 14)      ' Source Path:    0x00 (this packet does not contain a source path)
      buffer.SetValue(CByte(&H00), 15)      ' 0x00

      Dim checkSum As Byte = CalculateChecksum(buffer)
      buffer.SetValue(CByte(checkSum), 16)
      buffer.SetValue(CByte(&HF7), 17)

    Catch pEx As Exception

    End Try

    Return buffer

  End Function

  ''' <summary>
  ''' Set Zone Do Not Disturb
  ''' </summary>
  ''' <param name="cc"></param>
  ''' <param name="zz"></param>
  ''' <returns></returns>
  Protected Function SetZoneDoNotDisturbToggle(ByVal cc As Byte, ByVal zz As Byte) As Byte()

    Dim buffer(23) As Byte

    Try
      Dim kk As Byte = IIf(m_DeviceModel = "ACA-E5", &H71, &H70)

      buffer.SetValue(CByte(&HF0), 0)       ' F0 Start of Message
      buffer.SetValue(CByte(cc), 1)         ' Target Controller Id
      buffer.SetValue(CByte(&H00), 2)       ' Target Zone Id
      buffer.SetValue(CByte(&H7F), 3)       ' Target Keypad Number

      buffer.SetValue(CByte(cc), 4)         ' Source Controller Id
      buffer.SetValue(CByte(zz), 5)         ' Source Zone Id
      buffer.SetValue(CByte(kk), 6)         ' Source Keypad Number

      buffer.SetValue(PacketMessageType.Event, 7) ' Message Type (5 = Event)

      buffer.SetValue(CByte(&H05), 8)       ' 0x05
      buffer.SetValue(CByte(&H02), 9)       ' 0x02
      buffer.SetValue(CByte(&H00), 10)      ' 0x00
      buffer.SetValue(CByte(zz), 11)        ' Zone Number -1
      buffer.SetValue(CByte(&H00), 12)      ' 0x00
      buffer.SetValue(ZoneControlType.DoNotDisturb, 13) ' 0x06
      buffer.SetValue(CByte(&H00), 14)      ' 0x00
      buffer.SetValue(CByte(&H69), 15)      ' 0x69 Background Color Toggle
      buffer.SetValue(CByte(&H00), 16)      ' 0x00
      buffer.SetValue(CByte(&H00), 17)      ' 0x00
      buffer.SetValue(CByte(&H00), 18)      ' 0x00
      buffer.SetValue(CByte(&H07), 19)      ' 0x07
      buffer.SetValue(CByte(&H00), 20)      ' 0x00
      buffer.SetValue(CByte(&H01), 21)      ' 0x01

      Dim checkSum As Byte = CalculateChecksum(buffer)
      buffer.SetValue(CByte(checkSum), 22)
      buffer.SetValue(CByte(&HF7), 23)

    Catch pEx As Exception

    End Try

    Return buffer

  End Function

  ''' <summary>
  ''' Set Zone Do Not Disturb
  ''' </summary>
  ''' <param name="cc"></param>
  ''' <param name="zz"></param>
  ''' <param name="dnd"></param>
  ''' <returns></returns>
  Protected Function SetZoneDoNotDisturb(ByVal cc As Byte, ByVal zz As Byte, ByVal dnd As Byte) As Byte()

    Dim buffer(23) As Byte

    Try
      Dim kk As Byte = IIf(m_DeviceModel = "ACA-E5", &H71, &H70)

      buffer.SetValue(CByte(&HF0), 0)       ' F0 Start of Message
      buffer.SetValue(CByte(cc), 1)         ' Target Controller Id
      buffer.SetValue(CByte(&H00), 2)       ' Target Zone Id
      buffer.SetValue(CByte(&H7F), 3)       ' Target Keypad Number

      buffer.SetValue(CByte(cc), 4)         ' Source Controller Id
      buffer.SetValue(CByte(zz), 5)         ' Source Zone Id
      buffer.SetValue(CByte(kk), 6)         ' Source Keypad Number

      buffer.SetValue(PacketMessageType.SetData, 7) ' Message Type (0 = Set Data)

      buffer.SetValue(CByte(&H05), 8)       ' Target Path:  0x05 (Represented by the following 5 bytes)
      buffer.SetValue(CByte(&H02), 9)       ' Path Level 0: 0x02 (Root Menu)
      buffer.SetValue(CByte(&H00), 10)      ' Path Level 1: 0x00 (Run Menu)
      buffer.SetValue(CByte(zz), 11)        ' Path Level 2:   zz (Zone number -1)
      buffer.SetValue(CByte(&H00), 12)      ' Path Level 3: 0x00 (User Menu)
      buffer.SetValue(ZoneControlType.DoNotDisturb, 13)  ' Path Level 4: 0x06 (DoNotDisturb)
      buffer.SetValue(CByte(&H00), 14)      ' Source Path:  0x00 (This packet does not contain a source path)
      buffer.SetValue(CByte(&H00), 15)      ' 0x00
      buffer.SetValue(CByte(&H01), 16)      ' 0x00
      buffer.SetValue(CByte(&H01), 17)      ' 0x01
      buffer.SetValue(CByte(&H00), 18)      ' 0x00
      buffer.SetValue(CByte(&H01), 19)      ' 0x01
      buffer.SetValue(CByte(&H00), 20)      ' 0x00
      buffer.SetValue(CByte(dnd), 21)       ' dnd (0x00 = Off, 0x01 = On)

      Dim checkSum As Byte = CalculateChecksum(buffer)
      buffer.SetValue(CByte(checkSum), 22)
      buffer.SetValue(CByte(&HF7), 23)

    Catch pEx As Exception

    End Try

    Return buffer

  End Function

#End Region

#Region "Zone Party Mode"
  ''' <summary>
  ''' Set Zone Party Mode Ex
  ''' </summary>
  ''' <param name="cc"></param>
  ''' <param name="zz"></param>
  ''' <param name="mode"></param>
  Public Sub SetZonePartyModeEx(ByVal cc As Byte, ByVal zz As Byte, ByVal mode As Byte)
    AddCommand(SetZonePartyMode(cc, zz, mode))
    AddCommand(GetZoneStatus(cc, zz))
  End Sub

  ''' <summary>
  ''' Set Zone Party Mode Toggle Ex
  ''' </summary>
  ''' <param name="cc"></param>
  ''' <param name="zz"></param>
  Public Sub SetZonePartyModeToggleEx(ByVal cc As Byte, ByVal zz As Byte)
    AddCommand(SetZonePartyModeToggle(cc, zz))
  End Sub


  ''' <summary>
  ''' Get Zone Party Mode
  ''' </summary>
  ''' <param name="cc"></param>
  ''' <param name="zz"></param>
  ''' <returns></returns>
  Protected Function GetZonePartyMode(ByVal cc As Byte, ByVal zz As Byte) As Byte()

    Dim buffer(17) As Byte

    Try
      Dim kk As Byte = IIf(m_DeviceModel = "ACA-E5", &H71, &H70)

      buffer.SetValue(CByte(&HF0), 0)       ' F0 Start of Message
      buffer.SetValue(CByte(cc), 1)         ' Target Controller Id
      buffer.SetValue(CByte(&H00), 2)       ' Target Zone Id
      buffer.SetValue(CByte(&H7F), 3)       ' Target Keypad Number

      buffer.SetValue(CByte(cc), 4)         ' Source Controller Id
      buffer.SetValue(CByte(zz), 5)         ' Source Zone Id
      buffer.SetValue(CByte(kk), 6)         ' Source Keypad Number

      buffer.SetValue(PacketMessageType.RequestData, 7) ' Message Type

      buffer.SetValue(CByte(&H05), 8)       ' Target Path:    0x04 (Represented by the following 5 bytes)
      buffer.SetValue(CByte(&H02), 9)       ' Path Level 0:   0x02 (Root Menu)
      buffer.SetValue(CByte(&H00), 10)      ' Path Level 1:   0x00 (Run Menu)
      buffer.SetValue(CByte(zz), 11)        ' Path Level 2:   zz   (Zone number -1)
      buffer.SetValue(CByte(&H00), 12)      ' Path Level 3:   0x00 (User Menu)
      buffer.SetValue(CByte(&H07), 13)      ' Path Level 4:   0x07 (PARTY MODE)
      buffer.SetValue(CByte(&H00), 14)      ' Source Path:    0x00 (this packet does not contain a source path)
      buffer.SetValue(CByte(&H00), 15)      ' 0x00

      Dim checkSum As Byte = CalculateChecksum(buffer)
      buffer.SetValue(CByte(checkSum), 16)
      buffer.SetValue(CByte(&HF7), 17)

    Catch pEx As Exception

    End Try

    Return buffer

  End Function

  ''' <summary>
  ''' Set Zone Party Mode Toggle
  ''' </summary>
  ''' <param name="cc"></param>
  ''' <param name="zz"></param>
  ''' <returns></returns>
  Protected Function SetZonePartyModeToggle(ByVal cc As Byte, ByVal zz As Byte) As Byte()

    Dim buffer(23) As Byte

    Try
      Dim kk As Byte = IIf(m_DeviceModel = "ACA-E5", &H71, &H70)

      buffer.SetValue(CByte(&HF0), 0)       ' F0 Start of Message
      buffer.SetValue(CByte(cc), 1)         ' Target Controller Id
      buffer.SetValue(CByte(&H00), 2)       ' Target Zone Id
      buffer.SetValue(CByte(&H7F), 3)       ' Target Keypad Number

      buffer.SetValue(CByte(cc), 4)         ' Source Controller Id
      buffer.SetValue(CByte(zz), 5)         ' Source Zone Id
      buffer.SetValue(CByte(kk), 6)         ' Source Keypad Number

      buffer.SetValue(PacketMessageType.Event, 7) ' Message Type (5 = Event)

      buffer.SetValue(CByte(&H05), 8)       ' 0x05
      buffer.SetValue(CByte(&H02), 9)       ' 0x02
      buffer.SetValue(CByte(&H00), 10)      ' 0x00
      buffer.SetValue(CByte(zz), 11)        ' Zone Number -1
      buffer.SetValue(CByte(&H00), 12)      ' 0x00
      buffer.SetValue(ZoneControlType.PartyMode, 13) ' 0x07
      buffer.SetValue(CByte(&H00), 14)      ' 0x00
      buffer.SetValue(CByte(&H69), 15)      ' 0x69 Toggle
      buffer.SetValue(CByte(&H00), 16)      ' 0x00
      buffer.SetValue(CByte(&H00), 17)      ' 0x00
      buffer.SetValue(CByte(&H00), 18)      ' 0x00
      buffer.SetValue(CByte(&H08), 19)      ' 0x08
      buffer.SetValue(CByte(&H00), 20)      ' 0x00
      buffer.SetValue(CByte(&H01), 21)      ' 0x01

      Dim checkSum As Byte = CalculateChecksum(buffer)
      buffer.SetValue(CByte(checkSum), 22)
      buffer.SetValue(CByte(&HF7), 23)

    Catch pEx As Exception

    End Try

    Return buffer

  End Function

  ''' <summary>
  ''' Set Zone Party Mode
  ''' </summary>
  ''' <param name="cc"></param>
  ''' <param name="zz"></param>
  ''' <param name="mode"></param>
  ''' <returns></returns>
  Protected Function SetZonePartyMode(ByVal cc As Byte, ByVal zz As Byte, ByVal mode As Byte) As Byte()

    Dim buffer(23) As Byte

    Try
      Dim kk As Byte = IIf(m_DeviceModel = "ACA-E5", &H71, &H70)

      buffer.SetValue(CByte(&HF0), 0)       ' F0 Start of Message
      buffer.SetValue(CByte(cc), 1)         ' Target Controller Id
      buffer.SetValue(CByte(&H00), 2)       ' Target Zone Id
      buffer.SetValue(CByte(&H7F), 3)       ' Target Keypad Number

      buffer.SetValue(CByte(cc), 4)         ' Source Controller Id
      buffer.SetValue(CByte(zz), 5)         ' Source Zone Id
      buffer.SetValue(CByte(kk), 6)         ' Source Keypad Number

      buffer.SetValue(PacketMessageType.SetData, 7) ' Message Type (0 = Set Data)

      buffer.SetValue(CByte(&H05), 8)       ' Target Path:  0x05 (Represented by the following 5 bytes)
      buffer.SetValue(CByte(&H02), 9)       ' Path Level 0: 0x02 (Root Menu)
      buffer.SetValue(CByte(&H00), 10)      ' Path Level 1: 0x00 (Run Menu)
      buffer.SetValue(CByte(zz), 11)        ' Path Level 2:   zz (Zone number -1)
      buffer.SetValue(CByte(&H00), 12)      ' Path Level 3: 0x00 (User Menu)
      buffer.SetValue(ZoneControlType.PartyMode, 13)  ' Path Level 4: 0x07 (PartyMode)
      buffer.SetValue(CByte(&H00), 14)      ' Source Path:  0x00 (This packet does not contain a source path)
      buffer.SetValue(CByte(&H00), 15)      ' 0x00
      buffer.SetValue(CByte(&H01), 16)      ' 0x00
      buffer.SetValue(CByte(&H01), 17)      ' 0x01
      buffer.SetValue(CByte(&H00), 18)      ' 0x00
      buffer.SetValue(CByte(&H01), 19)      ' 0x01
      buffer.SetValue(CByte(&H00), 20)      ' 0x00
      buffer.SetValue(CByte(mode), 21)      ' mode (0x00 = Off, 0x01 = On, 0x02 = Master)

      Dim checkSum As Byte = CalculateChecksum(buffer)
      buffer.SetValue(CByte(checkSum), 22)
      buffer.SetValue(CByte(&HF7), 23)

    Catch pEx As Exception

    End Try

    Return buffer

  End Function

#End Region

#Region "Zone Status All Info"

  ''' <summary>
  ''' Get Zone State
  ''' </summary>
  ''' <param name="cc"></param>
  ''' <param name="zz"></param>
  ''' <returns></returns>
  Protected Function GetZoneStatus(ByVal cc As Byte, zz As Byte) As Byte()

    Dim buffer(16) As Byte

    Try
      Dim kk As Byte = IIf(m_DeviceModel = "ACA-E5", &H71, &H70)

      buffer.SetValue(CByte(&HF0), 0)       ' F0 Start of Message
      buffer.SetValue(CByte(cc), 1)         ' Target Controller Id
      buffer.SetValue(CByte(&H00), 2)       ' Target Zone Id
      buffer.SetValue(CByte(&H7F), 3)       ' Target Keypad Number 

      buffer.SetValue(CByte(cc), 4)         ' Source Controller Id
      buffer.SetValue(CByte(zz), 5)         ' Source Zone Id
      buffer.SetValue(CByte(kk), 6)         ' Source Keypad Number

      buffer.SetValue(PacketMessageType.RequestData, 7) ' Message Type (1 = RequestData)

      buffer.SetValue(CByte(&H04), 8)       ' Target Path:    0x04 (Represented by the following 4 bytes)
      buffer.SetValue(CByte(&H02), 9)       ' Path Level 0:   0x02 (Root Menu)
      buffer.SetValue(CByte(&H00), 10)      ' Path Level 1:   0x00 (Run Menu)
      buffer.SetValue(CByte(zz), 11)        ' Path Level 2:   zz   (Zone number -1)
      buffer.SetValue(CByte(&H07), 12)      ' Path Level 3:   0x07 (ZONE INFO)
      buffer.SetValue(CByte(&H00), 13)      ' Source Path:    0x00 (this packet does not contain a source path)
      buffer.SetValue(CByte(&H00), 14)

      Dim checkSum As Byte = CalculateChecksum(buffer)
      buffer.SetValue(CByte(checkSum), 15)
      buffer.SetValue(CByte(&HF7), 16)

    Catch pEx As Exception

    End Try

    Return buffer

  End Function

#End Region

#Region "Zone Keypad Events"

  ''' <summary>
  ''' UNO Keypad Events
  ''' </summary>
  ''' <param name="cc"></param>
  ''' <param name="zz"></param>
  ''' <param name="code"></param>
  Public Sub SendZoneKeypadControlEventEx(ByVal cc As Byte, ByVal zz As Byte, ByVal code As Byte)
    AddCommand(SendZoneKeypadControlEvent(cc, zz, code))
  End Sub

  ''' <summary>
  ''' UNO Keypad Events
  ''' </summary>
  ''' <param name="cc"></param>
  ''' <param name="zz"></param>
  ''' <returns></returns>
  Protected Function SendZoneKeypadControlEvent(ByVal cc As Byte, ByVal zz As Byte, ByVal code As Byte) As Byte()

    Dim buffer(21) As Byte

    Try
      Dim kk As Byte = IIf(m_DeviceModel = "ACA-E5", &H71, &H70)

      buffer.SetValue(CByte(&HF0), 0)       ' F0 Start of Message
      buffer.SetValue(CByte(cc), 1)         ' Target Controller Id
      buffer.SetValue(CByte(&H00), 2)       ' Target Zone Id
      buffer.SetValue(CByte(&H7F), 3)       ' Target Keypad Number

      buffer.SetValue(CByte(cc), 4)         ' Source Controller Id
      buffer.SetValue(CByte(zz), 5)         ' Source Zone Id
      buffer.SetValue(CByte(kk), 6)         ' Source Keypad Number

      buffer.SetValue(PacketMessageType.Event, 7) ' Message Type (5 = Event)

      buffer.SetValue(CByte(&H02), 8)       ' 0x02 Target Path:  0x02 (Represented by the following 2 bytes)
      buffer.SetValue(CByte(&H02), 9)       ' 0x02 Path Level 0: 0x02 (Root Menu)
      buffer.SetValue(CByte(&H00), 10)      ' 0x00 Path Level 1: 0x00 (Run Menu)
      buffer.SetValue(CByte(&H00), 11)      ' 0x00 Source Path:  0x00 (This packet does not contain a source path)
      buffer.SetValue(CByte(code), 12)      ' 0x00 Event ID Lo Byte:  Actual Keypad Command
      buffer.SetValue(CByte(&H00), 13)      ' 0x00 Evnet ID Hi Byte:  Not Used
      buffer.SetValue(CByte(&H00), 14)      ' 0x00 Unused
      buffer.SetValue(CByte(&H00), 15)      ' 0x00 Unused
      buffer.SetValue(CByte(&H00), 16)      ' 0x00 Unused
      buffer.SetValue(CByte(&H00), 17)      ' 0x00 Unused
      buffer.SetValue(CByte(&H00), 18)      ' 0x00 Unused
      buffer.SetValue(CByte(&H01), 19)      ' 0x01 Low Priority (does not require handshack)

      Dim checkSum As Byte = CalculateChecksum(buffer)
      buffer.SetValue(CByte(checkSum), 20)
      buffer.SetValue(CByte(&HF7), 21)

    Catch pEx As Exception

    End Try

    Return buffer

  End Function


#End Region

#Region "Zone Source Events"

  ''' <summary>
  ''' Source Control Events
  ''' </summary>
  ''' <param name="cc"></param>
  ''' <param name="zz"></param>
  ''' <param name="code"></param>
  Public Sub SendZoneSourceControlEventEx(ByVal cc As Byte, ByVal zz As Byte, ByVal code As Byte)
    AddCommand(SendZoneSourceControlEvent(cc, zz, code))
  End Sub

  ''' <summary>
  ''' Source Control Events
  ''' </summary>
  ''' <param name="cc"></param>
  ''' <param name="zz"></param>
  ''' <returns></returns>
  Protected Function SendZoneSourceControlEvent(ByVal cc As Byte, ByVal zz As Byte, ByVal code As Byte) As Byte()

    Dim buffer(21) As Byte

    Try
      Dim kk As Byte = IIf(m_DeviceModel = "ACA-E5", &H71, &H70)

      buffer.SetValue(CByte(&HF0), 0)       ' F0 Start of Message
      buffer.SetValue(CByte(cc), 1)         ' Target Controller Id
      buffer.SetValue(CByte(&H00), 2)       ' Target Zone Id
      buffer.SetValue(CByte(&H7F), 3)       ' Target Keypad Number

      buffer.SetValue(CByte(cc), 4)         ' Source Controller Id
      buffer.SetValue(CByte(zz), 5)         ' Source Zone Id
      buffer.SetValue(CByte(kk), 6)         ' Source Keypad Number

      buffer.SetValue(PacketMessageType.Event, 7) ' Message Type (5 = Event)

      buffer.SetValue(CByte(&H02), 8)       ' 0x02 Target Path:  0x02 (Represented by the following 2 bytes)
      buffer.SetValue(CByte(&H02), 9)       ' 0x02 Path Level 0: 0x02 (Root Menu)
      buffer.SetValue(CByte(&H00), 10)      ' 0x00 Path Level 1: 0x00 (Run Menu)
      buffer.SetValue(CByte(&H00), 11)      ' 0x00 Source Path:  0x00 (This packet does not contain a source path)
      buffer.SetValue(CByte(&HF1), 12)      ' 0xF1 Invert the next byte
      buffer.SetValue(CByte(&H40), 13)      ' 0x00 Event ID Lo Byte:  Remote Control Key Release
      buffer.SetValue(CByte(&H00), 14)      ' 0x00 Event ID Hi Byte:  Not Used
      buffer.SetValue(CByte(&H00), 15)      ' 0x00 Event Timestamp Lo Byte:  Unused
      buffer.SetValue(CByte(&H00), 16)      ' 0x00 Event Timestamp Hi Byte:  Unused
      buffer.SetValue(CByte(code), 17)      ' code Event Data Lo Byte:  Keycode
      buffer.SetValue(CByte(&H00), 18)      ' 0x00 Event Data Hi Byte:  Unused
      buffer.SetValue(CByte(&H01), 19)      ' 0x01 Low Priority (does not require handshack)

      Dim checkSum As Byte = CalculateChecksum(buffer)
      buffer.SetValue(CByte(checkSum), 20)
      buffer.SetValue(CByte(&HF7), 21)

    Catch pEx As Exception

    End Try

    Return buffer

  End Function

#End Region

#Region "Tuner Commands"

  ''' <summary>
  ''' Tuner Direct Preset Select
  ''' </summary>
  ''' <param name="cc"></param>
  ''' <param name="ss"></param>
  ''' <param name="category"></param>
  Public Sub TunerDirectCategorySelect(ByVal cc As Byte, ss As Byte, category As Byte)

    Try

      AddCommand(SetTunerCommand(cc, ss, TunerCommand.CATEGORY_SELECT, category))

    Catch pEx As Exception

    End Try

  End Sub

  ''' <summary>
  ''' Tuner Direct Preset Select
  ''' </summary>
  ''' <param name="cc"></param>
  ''' <param name="ss"></param>
  ''' <param name="frequency"></param>
  Public Sub TunerDirectBankSelect(ByVal cc As Byte, ss As Byte, frequency As String)

    Try

      SetTunerCommandEx(cc, ss, TunerCommand.DIRECT_BANK_SELECTION_MODE)
      For Each digit As Char In frequency
        Select Case digit
          Case "1"
            SetTunerCommandEx(cc, ss, TunerCommand.DIGIT_1)
          Case "2"
            SetTunerCommandEx(cc, ss, TunerCommand.DIGIT_2)
          Case "3"
            SetTunerCommandEx(cc, ss, TunerCommand.DIGIT_3)
          Case "4"
            SetTunerCommandEx(cc, ss, TunerCommand.DIGIT_4)
          Case "5"
            SetTunerCommandEx(cc, ss, TunerCommand.DIGIT_5)
          Case "6"
            SetTunerCommandEx(cc, ss, TunerCommand.DIGIT_6)
        End Select
      Next

    Catch pEx As Exception

    End Try

  End Sub

  ''' <summary>
  ''' Tuner Direct Preset Select
  ''' </summary>
  ''' <param name="cc"></param>
  ''' <param name="ss"></param>
  ''' <param name="frequency"></param>
  Public Sub TunerDirectPresetSelect(ByVal cc As Byte, ss As Byte, frequency As String)

    Try

      SetTunerCommandEx(cc, ss, TunerCommand.DIRECT_PRESET_SELECTION_MODE)
      For Each digit As Char In frequency
        Select Case digit
          Case "1"
            SetTunerCommandEx(cc, ss, TunerCommand.DIGIT_1)
          Case "2"
            SetTunerCommandEx(cc, ss, TunerCommand.DIGIT_2)
          Case "3"
            SetTunerCommandEx(cc, ss, TunerCommand.DIGIT_3)
          Case "4"
            SetTunerCommandEx(cc, ss, TunerCommand.DIGIT_4)
          Case "5"
            SetTunerCommandEx(cc, ss, TunerCommand.DIGIT_5)
          Case "6"
            SetTunerCommandEx(cc, ss, TunerCommand.DIGIT_6)
        End Select
      Next

    Catch pEx As Exception

    End Try

  End Sub

  ''' <summary>
  ''' Tuner Direct Tuning Mode
  ''' </summary>
  ''' <param name="cc"></param>
  ''' <param name="ss"></param>
  ''' <param name="frequency"></param>
  Public Sub TunerDirectTuneMode(ByVal cc As Byte, ss As Byte, frequency As String)

    Try

      SetTunerCommandEx(cc, ss, TunerCommand.DIRECT_TUNING_MODE)
      frequency = frequency.PadLeft(3, "0")
      For Each digit As Char In frequency
        Select Case digit
          Case "0"
            SetTunerCommandEx(cc, ss, TunerCommand.DIGIT_0)
          Case "1"
            SetTunerCommandEx(cc, ss, TunerCommand.DIGIT_1)
          Case "2"
            SetTunerCommandEx(cc, ss, TunerCommand.DIGIT_2)
          Case "3"
            SetTunerCommandEx(cc, ss, TunerCommand.DIGIT_3)
          Case "4"
            SetTunerCommandEx(cc, ss, TunerCommand.DIGIT_4)
          Case "5"
            SetTunerCommandEx(cc, ss, TunerCommand.DIGIT_5)
          Case "6"
            SetTunerCommandEx(cc, ss, TunerCommand.DIGIT_6)
          Case "7"
            SetTunerCommandEx(cc, ss, TunerCommand.DIGIT_7)
          Case "8"
            SetTunerCommandEx(cc, ss, TunerCommand.DIGIT_8)
          Case "9"
            SetTunerCommandEx(cc, ss, TunerCommand.DIGIT_9)
        End Select
      Next

    Catch pEx As Exception

    End Try

  End Sub

  ''' <summary>
  ''' Set Zone Power State Ex
  ''' </summary>
  ''' <param name="cc"></param>
  ''' <param name="ss"></param>
  ''' <param name="cmd"></param>
  Public Sub SetTunerCommandEx(ByVal cc As Byte, ss As Byte, cmd As Byte)
    AddCommand(SetTunerCommand(cc, ss, cmd), True)
  End Sub

  ''' <summary>
  ''' Send Tuner Command
  ''' </summary>
  ''' <param name="cc"></param>
  ''' <param name="ss"></param>
  ''' <param name="cmd"></param>
  ''' <returns></returns>
  Protected Function SetTunerCommand(ByVal cc As Byte, ss As Byte, cmd As Byte, Optional cmd2 As Byte = &H70) As Byte()

    Dim buffer(22) As Byte

    Try

      buffer.SetValue(CByte(&HF0), 0)             ' F0 Start of Message
      buffer.SetValue(CByte(cc), 1)               ' Target Controller Id:     Id assigned to the controller at startup
      buffer.SetValue(CByte(&H7D), 2)             ' Target Zone Id:           0x7D is the RNET Peripheral
      buffer.SetValue(CByte(ss), 3)               ' Source Number:            Zero based, 0 = Source 1

      buffer.SetValue(CByte(&H00), 4)             ' Source Controller Id:    For 3rd party devices, this should be set to &H00
      buffer.SetValue(CByte(&H00), 5)             ' Source Zone (Port) Id:   For 3rd party devices, this should be set to a unique value
      buffer.SetValue(CByte(&H71), 6)             ' Third party Id:          &H71

      buffer.SetValue(PacketMessageType.Event, 7) ' Message Type (5 = Event)

      buffer.SetValue(CByte(&H02), 8)             ' Target Path:            Represented by the following 2 bytes
      buffer.SetValue(CByte(&H01), 9)             ' Path Level 0:           StdInterface
      buffer.SetValue(CByte(&H00), 10)            ' Path Level 1:           EventHandler
      buffer.SetValue(CByte(&H02), 11)            ' Source Path:            Represented by the following 2 bytes
      buffer.SetValue(CByte(&H01), 12)            ' Path Level 0:           StdInterface  
      buffer.SetValue(CByte(&H00), 13)            ' Path Level 1:           EventHandler
      buffer.SetValue(CByte(cmd), 14)             ' Event Hi Byte:          Tuner command
      buffer.SetValue(CByte(&H00), 15)            ' Event Lo Byte:          Unused        
      buffer.SetValue(CByte(cmd2), 16)            ' Not specified:          &H70 or XM Category when used with Select Category
      buffer.SetValue(CByte(&H00), 17)            ' Not specified:          Unknown
      buffer.SetValue(CByte(ss), 18)              ' Source Number:          Must be the same as byte 3
      buffer.SetValue(CByte(&H00), 19)            ' Not specified:          Unknown
      buffer.SetValue(CByte(&H01), 20)            ' Not specified:          Unknown

      Dim checkSum As Byte = CalculateChecksum(buffer)
      buffer.SetValue(CByte(checkSum), 21)
      buffer.SetValue(CByte(&HF7), 22)

    Catch pEx As Exception

    End Try

    Return buffer

  End Function

#End Region

#End Region

#Region "Audio Device Serial Support"

  ''' <summary>
  ''' Event handler for Com Port data received
  ''' </summary>
  ''' <param name="sender"></param>
  ''' <param name="e"></param>
  ''' <remarks></remarks>
  Private Sub DataReceived(ByVal sender As Object,
                           ByVal e As System.IO.Ports.SerialDataReceivedEventArgs) _
                           Handles serialPort.DataReceived

    Dim strMessage As String = "Entered Serial DataReceived() function."
    Call WriteMessage(strMessage, MessageType.Debug)

    '
    ' Read from com port and buffer till we get a vbLf
    '
    Try

      Dim Packet(0) As Byte

      Do While serialPort.BytesToRead > 0

        Do While serialPort.BytesToRead > 0
          Dim DataByte As Byte = serialPort.ReadByte
          Select Case DataByte
            Case &HF0 ' Start of Message
              ReDim Packet(0)
              Packet(0) = DataByte

            Case &HF7 ' End of Message
              Dim packetLength = Packet.Length
              ReDim Preserve Packet(packetLength)
              Packet(packetLength) = DataByte

              ProcessReceived(Packet)
              ReDim Packet(0)

            Case Else ' Normal Packet
              Dim packetLength = Packet.Length
              ReDim Preserve Packet(packetLength)
              Packet(packetLength) = DataByte

          End Select
          Thread.Sleep(0)
        Loop

        Thread.Sleep(25)
      Loop

    Catch pEx As Exception
      WriteMessage(pEx.ToString, MessageType.Error)
    End Try

  End Sub

#End Region

#Region "Ethernet Support"

  Dim TcpClient As System.Net.Sockets.TcpClient
  Dim NetworkStream As System.Net.Sockets.NetworkStream
  Dim ReadThread As Threading.Thread

  ''' <summary>
  ''' Establish connection to Ethernet Module
  ''' </summary>
  ''' <param name="Ip"></param>
  ''' <param name="Port"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function ConnectToEthernet(ByVal Ip As String, ByVal Port As Integer) As Boolean

    Dim strMessage As String
    Dim IPAddress As String = ResolveAddress(Ip)

    Try

      Try
        '
        ' Create TCPClient
        '
        TcpClient = New TcpClient(IPAddress, Port)

      Catch pEx As SocketException
        '
        ' Process Exception
        '
        strMessage = String.Format("Ethernet connection could not be made to {0} ({1}:{2}) - {3}",
                                  IPAddress, Ip.ToString, Port.ToString, pEx.Message)
        Call WriteMessage(strMessage, MessageType.Debug)
        Return False
      End Try

      NetworkStream = TcpClient.GetStream()
      ReadThread = New Thread(New ThreadStart(AddressOf EthernetReadThreadProc))
      ReadThread.Name = "EthernetReadThreadProc"
      ReadThread.Start()

    Catch pEx As Exception
      '
      ' Process Exception
      '
      Call ProcessError(pEx, "ConnectToEthernet()")
      Return False
    End Try

    Return True

  End Function

  ''' <summary>
  ''' Disconnection From Ethernet Module
  ''' </summary>
  ''' <remarks></remarks>
  Public Sub DisconnectEthernet()

    Try
      If ReadThread.IsAlive = True Then
        ReadThread.Abort()
      End If
      NetworkStream.Close()
      TcpClient.Close()
    Catch ex As Exception
      '
      ' Ignore Exception
      '
    End Try

  End Sub

  ''' <summary>
  ''' Send Message to connected IP address (first send buffer length and then the buffer holding message)
  ''' </summary>
  ''' <param name="Packet"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Protected Function SendMessageToEthernet(ByVal Packet() As Byte) As Boolean

    Try

      If TcpClient.Connected = True Then
        NetworkStream.Write(Packet, 0, Packet.Length)
        Return True
      Else
        Call WriteMessage("Attempted to write to a closed Ethernet stream in SendMessageToEthernet()", MessageType.Warning)
        Return False
      End If

    Catch pEx As Exception
      Call WriteMessage("Attempted to write to a closed Ethernet stream in SendMessageToEthernet()", MessageType.Warning)
      Return False
    End Try

  End Function

  ''' <summary>
  ''' Process to Read Data From TCP Client
  ''' </summary>
  ''' <remarks></remarks>
  Protected Sub EthernetReadThreadProc()

    Try

      Dim r As New BinaryReader(NetworkStream)

      Dim strMessage As String = "Entered EthernetReadThreadProc() subroutine."
      WriteMessage(strMessage, MessageType.Debug)

      Dim Packet(0) As Byte

      '
      ' Stay in EthernetReadThreadProc while client is connected
      '
      Do While TcpClient.Connected = True

        Do While NetworkStream.DataAvailable = True

          Dim DataByte As Byte = r.ReadByte()
          Select Case DataByte
            Case &HF0 ' Start of Message
              ReDim Packet(0)
              Packet(0) = DataByte

            Case &HF7 ' End of Message
              Dim packetLength = Packet.Length
              ReDim Preserve Packet(packetLength)
              Packet(packetLength) = DataByte

              ProcessReceived(Packet)
              ReDim Packet(0)

            Case Else ' Normal Packet
              Dim packetLength = Packet.Length
              ReDim Preserve Packet(packetLength)
              Packet(packetLength) = DataByte

          End Select

          Thread.Sleep(0)
        Loop

        Thread.Sleep(25)
      Loop

    Catch ab As ThreadAbortException
      '
      ' Process Thread Abort Exception
      '
      Call WriteMessage("Abort requested on EthernetReadThreadProc", MessageType.Debug)
      Return
    Catch pEx As Exception
      '
      ' Process Exception
      '
      Call ProcessError(pEx, "EthernetReadThreadProc()")
    Finally
      '
      ' Indicate we are no longer connected to the Audio Device
      '
      gDeviceConnected = False
    End Try

  End Sub

  ''' <summary>
  ''' Check ip string to be an ip address or if not try to resolve using DNS
  ''' </summary>
  ''' <param name="hostNameOrAddress"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function ResolveAddress(ByVal hostNameOrAddress As String) As String

    Try
      '
      ' Attempt to identify fqdn as an IP address
      '
      IPAddress.Parse(hostNameOrAddress)

      '
      ' If this did not throw then it is a valid IP address
      '
      Return hostNameOrAddress
    Catch ex As Exception
      Try
        ' Try to resolve it through DNS if it is not in IP address form
        ' and use the first IP address if defined as round robbin in DNS
        Dim ipAddress As IPAddress = Dns.GetHostEntry(hostNameOrAddress).AddressList(0)

        Return ipAddress.ToString
      Catch pEx As Exception
        Return ""
      End Try

    End Try

  End Function

#End Region

#Region "HomeSeer Device Support"

  ''' <summary>
  ''' Updates the Audio Zone Device
  ''' </summary>
  ''' <param name="ss"></param>
  ''' <param name="keypadMessage"></param>
  Private Sub SetKeypadSourceMessage(ss As Byte, ByVal keypadMessage As String)

    Try

      For Each AudioZone As hspi_audio_zone In gAudioZones.FindAll(Function(s) s.Source = ss)
        Dim audioZoneId As String = String.Format("{0}.{1}.{2}", Me.DeviceId, AudioZone.ControllerId, AudioZone.ZoneId)
        Dim dv_addr As String = String.Format("Russound{0}-{1}", audioZoneId, "audiozone-source")

        SetKeypadMessage(AudioZone.ControllerId, AudioZone.ZoneId, keypadMessage)
      Next

    Catch pEx As Exception
      Call ProcessError(pEx, "SetKeypadSourceMessage()")
    End Try

  End Sub

  ''' <summary>
  ''' Updates the Audio Zone Device
  ''' </summary>
  ''' <param name="cc"></param>
  ''' <param name="zz"></param>
  ''' <param name="keypadMessage"></param>
  Private Sub SetKeypadMessage(ByVal cc As Byte, zz As Byte, ByVal keypadMessage As String)

    Dim audioZoneId As String = String.Format("{0}.{1}.{2}", Me.DeviceId, cc, zz)
    Dim dv_addr As String = String.Format("Russound{0}-{1}", audioZoneId, "audiozone-keypad-display")

    Try
      hspi_devices.SetDeviceString(dv_addr, keypadMessage)
    Catch pEx As Exception
      Call ProcessError(pEx, "SetKeypadSourceMessage()")
    End Try

  End Sub

  ''' <summary>
  ''' Updates the Audio Zone Device
  ''' </summary>
  ''' <param name="cc"></param>
  ''' <param name="zz"></param>
  ''' <param name="zone_key"></param>
  ''' <param name="zone_value"></param>
  Private Sub SetAudioZoneValue(ByVal cc As Byte, zz As Byte, ByVal zone_key As String, ByVal zone_value As Byte)

    Dim audioZoneId As String = String.Format("{0}.{1}.{2}", Me.DeviceId, cc, zz)
    Dim dv_addr As String = String.Format("Russound{0}-{1}", audioZoneId, zone_key)

    Try
      gAudioZones.Find(Function(s) s.DeviceAddr = dv_addr).SetPropertyValue(zone_key, zone_value)
      hspi_devices.SetDeviceValue(dv_addr, zone_value)
    Catch pEx As Exception
      Call ProcessError(pEx, "SetAudioZoneValue()")
    End Try

  End Sub

  ''' <summary>
  ''' Gets the Key Names
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetAudioZoneKeyTypes() As String()

    Try

      Dim Keys() As String = {"Zone Master Controls",
                              "Zone Audio Controls",
                              "Zone Keypads",
                              "Zone Misc",
                              "ST2 Smart Tuner"}

      Return Keys

    Catch pEx As Exception
      Return Nothing
    End Try

  End Function

  ''' <summary>
  ''' Gets the Key Names
  ''' </summary>
  ''' <param name="strKeyType"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetAudioZoneKeyNames(ByVal strKeyType As String) As String()

    Try

      Select Case strKeyType
        Case "Zone Master Controls"
          Dim Keys() As String = {"audiozone-power",
                                  "audiozone-source",
                                  "audiozone-source-controls",
                                  "audiozone-partymode",
                                  "audiozone-dnd",
                                  "audiozone-turnon-vol"}
          Return Keys

        Case "Zone Audio Controls"
          Dim Keys() As String = {"audiozone-volume",
                                  "audiozone-bass",
                                  "audiozone-treble",
                                  "audiozone-balance",
                                  "audiozone-loudness"}
          Return Keys

        Case "Zone Keypads"
          Dim Keys() As String = {"audiozone-keypad-controls",
                                  "audiozone-keypad-bg-color",
                                  "audiozone-keypad-display"}
          Return Keys

        Case "Zone Misc"
          Dim Keys() As String = {"audiozone-system-onstate",
                                  "audiozone-source-shared"}

          Return Keys

        Case "ST2 Smart Tuner"
          Dim Keys() As String = {"tuner-power",
                                  "tuner-preset",
                                  "tuner-bank",
                                  "tuner-radio-frequency",
                                  "tuner-radio-controls",
                                  "tuner-xm-channel",
                                  "tuner-xm-category",
                                  "tuner-xm-category-channel"}

          Return Keys

        Case Else
          Return Nothing
      End Select

    Catch pEx As Exception
      Return Nothing
    End Try

  End Function

  ''' <summary>
  ''' Gets the Key Friendly Name
  ''' </summary>
  ''' <param name="strKey"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetKeyFriendlyName(ByVal strKey As String) As String

    Select Case strKey
      ' Master Controls
      Case "audiozone-power" : Return "Power"
      Case "audiozone-source" : Return "Source"
      Case "audiozone-source-controls" : Return "Source Controls"
      Case "audiozone-partymode" : Return "Party Mode"
      Case "audiozone-dnd" : Return "Do Not Distrub"
      Case "audiozone-turnon-vol" : Return "Turn On Volume"

      ' Zone Audio Controls
      Case "audiozone-volume" : Return "Volume"
      Case "audiozone-bass" : Return "Bass"
      Case "audiozone-treble" : Return "Treble"
      Case "audiozone-balance" : Return "Balance"
      Case "audiozone-loudness" : Return "Loudness"

      ' Zone Keypads
      Case "audiozone-keypad-controls" : Return "Keypad Controls"
      Case "audiozone-keypad-bg-color" : Return "Keypad Background Color"
      Case "audiozone-keypad-display" : Return "Keypad Display Messages"

      ' Zone Misc
      Case "audiozone-system-onstate" : Return "System OnState"
      Case "audiozone-source-shared" : Return "Shared Source"

      ' ST2 Smart Tuner
      Case "tuner-power" : Return "Power"
      Case "tuner-preset" : Return "Preset"
      Case "tuner-bank" : Return "Bank"
      Case "tuner-radio-frequency" : Return "Radio Frequency"
      Case "tuner-radio-controls" : Return "Radio Controls"
      Case "tuner-xm-channel" : Return "XM Channel"
      Case "tuner-xm-category" : Return "XM Category"
      Case "tuner-xm-category-channel" : Return "XM Category Channel"

      Case Else : Return strKey
    End Select

  End Function

  ''' <summary>
  ''' Gets the Key Name from Key Type
  ''' </summary>
  ''' <param name="strKey"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Function GetKeyType(ByVal strKey As String) As String

    Select Case strKey
      ' Master Controls
      Case "audiozone-power" : Return "Zone Master Controls"
      Case "audiozone-source" : Return "Zone Master Controls"
      Case "audiozone-source-controls" : Return "Zone Master Controls"
      Case "audiozone-partymode" : Return "Zone Master Controls"
      Case "audiozone-dnd" : Return "Zone Master Controls"
      Case "audiozone-turnon-vol" : Return "Zone Master Controls"

      ' Zone Audio Controls
      Case "audiozone-volume" : Return "Zone Audio Controls"
      Case "audiozone-bass" : Return "Zone Audio Controls"
      Case "audiozone-treble" : Return "Zone Audio Controls"
      Case "audiozone-balance" : Return "Zone Audio Controls"
      Case "audiozone-loudness" : Return "Zone Audio Controls"

      ' Zone Keypads
      Case "audiozone-keypad-controls" : Return "Zone Keypads"
      Case "audiozone-keypad-bg-color" : Return "Zone Keypads"
      Case "audiozone-keypad-display" : Return "Zone Keypads"

      ' Zone Misc
      Case "audiozone-system-onstate" : Return "Zone Misc"
      Case "audiozone-source-shared" : Return "Zone Misc"

      ' ST2 Smart Tuner
      Case "tuner-power" : Return "ST2 Smart Tuner"
      Case "tuner-preset" : Return "ST2 Smart Tuner"
      Case "tuner-bank" : Return "ST2 Smart Tuner"
      Case "tuner-radio-frequency" : Return "ST2 Smart Tuner"
      Case "tuner-radio-controls" : Return "ST2 Smart Tuner"
      Case "tuner-xm-channel" : Return "ST2 Smart Tuner"
      Case "tuner-xm-category" : Return "ST2 Smart Tuner"
      Case "tuner-xm-category-channel" : Return "ST2 Smart Tuner"

      Case Else : Return "Unknown Device"
    End Select

  End Function

  ''' <summary>
  ''' Creates the Audio Root Device
  ''' </summary>
  ''' <param name="strDeviceId"></param>
  ''' <param name="strDeviceDevice"></param>
  ''' <param name="strDeviceType"></param>
  ''' <param name="dv_ref_child"></param>
  ''' <returns></returns>
  Public Function CreateAudioRootDevice(ByVal strDeviceId As String,
                                        ByVal strDeviceDevice As String,
                                        ByVal strDeviceType As String,
                                        ByVal dv_ref_child As Integer) As Integer

    Dim dv As Scheduler.Classes.DeviceClass

    Dim dv_ref As Integer = 0
    Dim dv_misc As Integer = 0

    Dim dv_name As String = ""
    Dim dv_type As String = ""
    Dim dv_addr As String = ""

    Dim DeviceShowValues As Boolean = False

    Try
      '
      ' Set the local variables
      '
      dv_name = String.Format("Russound{0} {1} Root", strDeviceId, strDeviceType)
      dv_addr = String.Format("Russound{0}_{1}-root", strDeviceId, strDeviceType.Replace(" ", "-"))
      dv_type = String.Format("Russound AudioZone {0} Root", strDeviceType)

      dv = LocateDeviceByAddr(dv_addr)
      Dim bDeviceExists As Boolean = Not dv Is Nothing

      If bDeviceExists = True Then
        '
        ' Lets use the existing device
        '
        dv_addr = dv.Address(hs)
        dv_ref = dv.Ref(hs)

        Call WriteMessage(String.Format("Updating existing HomeSeer {0} root device.", dv_name), MessageType.Debug)

      Else
        '
        ' Create A HomeSeer Device
        '
        dv_ref = hs.NewDeviceRef(dv_name)
        If dv_ref > 0 Then
          dv = hs.GetDeviceByRef(dv_ref)
        End If

        Call WriteMessage(String.Format("Creating New HomeSeer {0} root device.", dv_name), MessageType.Debug)

      End If

      '
      ' Define the HomeSeer device
      '
      dv.Address(hs) = dv_addr
      dv.Interface(hs) = IFACE_NAME
      dv.InterfaceInstance(hs) = Instance

      '
      ' Update location properties on new devices only
      '
      If bDeviceExists = False Then
        dv.Location(hs) = IFACE_NAME & " Plugin"
        dv.Location2(hs) = String.Format("Russound{0} {1}", strDeviceId, strDeviceType)
      End If

      '
      ' The following simply shows up in the device properties but has no other use
      '
      dv.Device_Type_String(hs) = dv_type

      '
      ' Set the DeviceTypeInfo
      '
      Dim DT As New DeviceTypeInfo
      DT.Device_API = DeviceTypeInfo.eDeviceAPI.Plug_In
      DT.Device_Type = DeviceTypeInfo.eDeviceType_Plugin.Root
      dv.DeviceType_Set(hs) = DT

      '
      ' Make this a parent root device
      '
      dv.Relationship(hs) = Enums.eRelationship.Parent_Root
      dv.AssociatedDevice_Add(hs, dv_ref_child)

      Dim image As String = "device_root.png"

      Dim VSPair As VSPair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
      VSPair.PairType = VSVGPairType.SingleValue
      VSPair.Value = 0
      VSPair.Status = "Root"
      VSPair.Render = Enums.CAPIControlType.Values
      hs.DeviceVSP_AddPair(dv_ref, VSPair)

      Dim VGPair As VGPair = New VGPair()
      VGPair.PairType = VSVGPairType.SingleValue
      VGPair.Set_Value = 0
      VGPair.Graphic = String.Format("{0}{1}", gImageDir, image)
      hs.DeviceVGP_AddPair(dv_ref, VGPair)

      '
      ' Update the Device Misc Bits
      '
      If DeviceShowValues = True Then
        dv.MISC_Set(hs, Enums.dvMISC.SHOW_VALUES)
      End If

      If bDeviceExists = False Then
        dv.MISC_Set(hs, Enums.dvMISC.NO_LOG)
      End If

      dv.Status_Support(hs) = False

      hs.SaveEventsDevices()

    Catch pEx As Exception

    End Try

    Return dv_ref

  End Function

  ''' <summary>
  ''' Function to create our plug-in connection device used for status and control
  ''' </summary>
  ''' <param name="device_conn"></param>
  ''' <param name="dev_code"></param>
  ''' <param name="device_connected"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function UpdateAudioConnectionDevice(ByVal device_conn As String, ByVal dev_code As Integer, ByVal device_connected As Boolean) As String

    Dim dv As Scheduler.Classes.DeviceClass

    Dim dv_ref As Integer = 0
    Dim dv_misc As Integer = 0

    Dim dv_name As String = ""
    Dim dv_type As String = ""
    Dim dv_addr As String = ""

    Try
      '
      ' Set the local variables
      '
      Dim strDeviceType As String = "Controller"
      dv_type = "Russound AudioZone Controller"
      dv_name = String.Format("Russound{0} {1}", dev_code, "Connection")
      dv_addr = String.Format("Russound{0}_Connection", dev_code.ToString)
      dv = LocateDeviceByAddr(dv_addr)

      Dim bDeviceExists As Boolean = Not dv Is Nothing

      If bDeviceExists = True Then
        '
        ' Lets use the existing device
        '
        dv_addr = dv.Address(hs)
        dv_ref = dv.Ref(hs)

        Call WriteMessage(String.Format("Updating existing HomeSeer {0} device.", dv_name), MessageType.Debug)

      Else
        '
        ' Create A HomeSeer Device
        '
        dv_ref = hs.NewDeviceRef(dv_name)
        If dv_ref > 0 Then
          dv = hs.GetDeviceByRef(dv_ref)
        End If

        Call WriteMessage(String.Format("Creating New HomeSeer {0} device.", dv_name), MessageType.Debug)

      End If

      '
      ' Store the UUID for the device
      '
      Dim pdata As New clsPlugExtraData
      pdata.AddNamed("UUID", DeviceId)
      dv.PlugExtraData_Set(hs) = pdata

      '
      ' Define the HomeSeer device
      '
      dv.Address(hs) = dv_addr
      dv.Interface(hs) = IFACE_NAME
      dv.InterfaceInstance(hs) = Instance

      '
      ' Update location properties on new devices only
      '
      If bDeviceExists = False Then
        dv.Location(hs) = IFACE_NAME & " Plugin"
        dv.Location2(hs) = String.Format("Russound{0} {1}", dev_code, strDeviceType)
      End If

      '
      ' The following simply shows up in the device properties but has no other use
      '
      dv.Device_Type_String(hs) = dv_type

      '
      ' Set the DeviceTypeInfo
      '
      Dim DT As New DeviceTypeInfo
      DT.Device_API = DeviceTypeInfo.eDeviceAPI.Plug_In
      DT.Device_Type = DeviceTypeInfo.eDeviceType_Plugin.Root
      dv.DeviceType_Set(hs) = DT

      '
      ' Make this device a child of the root
      '
      dv.AssociatedDevice_ClearAll(hs)
      Dim dvp_ref As Integer = CreateAudioRootDevice(dev_code.ToString, "Connection", strDeviceType, dv_ref)
      If dvp_ref > 0 Then
        dv.AssociatedDevice_Add(hs, dvp_ref)
      End If
      dv.Relationship(hs) = Enums.eRelationship.Child

      '
      ' Update the last change date
      ' 
      dv.Last_Change(hs) = DateTime.Now

      Dim VSPair As VSPair

      VSPair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
      VSPair.PairType = VSVGPairType.SingleValue
      VSPair.Value = -3
      VSPair.Status = ""
      VSPair.Render = Enums.CAPIControlType.Values
      hs.DeviceVSP_AddPair(dv_ref, VSPair)

      VSPair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
      VSPair.PairType = VSVGPairType.SingleValue
      VSPair.Value = -2
      VSPair.Status = "Disconnect"
      VSPair.Render = Enums.CAPIControlType.Values
      hs.DeviceVSP_AddPair(dv_ref, VSPair)

      VSPair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
      VSPair.PairType = VSVGPairType.SingleValue
      VSPair.Value = -1
      VSPair.Status = "Reconnect"
      VSPair.Render = Enums.CAPIControlType.Values
      hs.DeviceVSP_AddPair(dv_ref, VSPair)

      VSPair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
      VSPair.PairType = VSVGPairType.SingleValue
      VSPair.Value = 0
      VSPair.Status = "Disconnected"
      VSPair.Render = Enums.CAPIControlType.Values
      hs.DeviceVSP_AddPair(dv_ref, VSPair)

      VSPair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
      VSPair.PairType = VSVGPairType.SingleValue
      VSPair.Value = 1
      VSPair.Status = "Connected"
      VSPair.Render = Enums.CAPIControlType.Values
      hs.DeviceVSP_AddPair(dv_ref, VSPair)

      Dim VGPair As VGPair

      '
      ' Add VGPairs
      '
      VGPair = New VGPair()
      VGPair.PairType = VSVGPairType.Range
      VGPair.RangeStart = -3
      VGPair.RangeEnd = 0
      VGPair.Graphic = String.Format("{0}{1}", gImageDir, "network_disconnected.png")
      hs.DeviceVGP_AddPair(dv_ref, VGPair)

      VGPair = New VGPair()
      VGPair.PairType = VSVGPairType.SingleValue
      VGPair.Set_Value = 1
      VGPair.Graphic = String.Format("{0}{1}", gImageDir, "network_connected.png")
      hs.DeviceVGP_AddPair(dv_ref, VGPair)

      '
      ' Update the Device Misc Bits
      '
      dv.MISC_Set(hs, Enums.dvMISC.SHOW_VALUES)
      If bDeviceExists = False Then
        dv.MISC_Set(hs, Enums.dvMISC.NO_LOG)
      End If

      dv.Status_Support(hs) = False

      hs.SaveEventsDevices()

      '
      ' Update the connection status
      '
      Dim dv_value As Long = IIf(device_connected = True, 1, 0)
      hspi_devices.SetDeviceValue(dv_addr, dv_value)

      Return ""

    Catch pEx As Exception
      '
      ' Process program exception
      '
      Call ProcessError(pEx, "UpdateAudioConnectionDevice()")
      Return pEx.ToString
    End Try

  End Function

  ''' <summary>
  ''' 
  ''' </summary>
  ''' <param name="dv_addr"></param>
  ''' <param name="bUpdateOnly"></param>
  ''' <returns></returns>
  Public Function CreateAudioZoneDevice(ByVal dv_addr As String, ByVal bUpdateOnly As Boolean)

    Dim dv As Scheduler.Classes.DeviceClass

    Dim dv_ref As Integer = 0
    Dim dv_misc As Integer = 0

    Dim dv_name As String = String.Empty
    Dim dv_type As String = String.Empty

    Dim DeviceShowValues As Boolean = True
    Dim DevicePairs As New ArrayList

    Try

      '
      ' Set the local variables
      '
      Dim regexPattern As String = "Russound(?<dd>(\d+))\.(?<cc>(\d+))\.(?<zz>(\d+))-(?<type>(.+))"

      Dim device_id As String = Regex.Match(dv_addr, regexPattern).Groups("dd").ToString()
      Dim cc As String = Regex.Match(dv_addr, regexPattern).Groups("cc").ToString()
      Dim zz As String = Regex.Match(dv_addr, regexPattern).Groups("zz").ToString()
      Dim device_type As String = Regex.Match(dv_addr, regexPattern).Groups("type").ToString()

      Dim zone_id As String = String.Format("{0}.{1}.{2}", device_id, cc, zz)
      Dim zone_number As Integer = 0
      Select Case Me.DeviceModel
        Case "CAS44"
          zone_number = cc * 4 + (zz + 1)
        Case Else
          zone_number = cc * 6 + (zz + 1)
      End Select
      Dim zone_name As String = String.Format("Zone{0}", zone_number.ToString.PadLeft(2, "0"))
      Dim device_desc As String = GetKeyFriendlyName(device_type)

      dv_type = String.Format("Russound AudioZone {0}", device_desc)
      dv_name = String.Format("Russound{0} {1} {2}", device_id, zone_name, device_desc)

      dv = LocateDeviceByAddr(dv_addr)

      Dim bDeviceExists As Boolean = Not dv Is Nothing
      If bUpdateOnly = True And bDeviceExists = False Then Return ""

      If bDeviceExists = True Then
        '
        ' Lets use the existing device
        '
        dv_addr = dv.Address(hs)
        dv_ref = dv.Ref(hs)

        Call WriteMessage(String.Format("Updating existing HomeSeer {0} device.", dv_name), MessageType.Debug)

      Else
        '
        ' Create A HomeSeer Device
        '
        dv_ref = hs.NewDeviceRef(dv_name)
        If dv_ref > 0 Then
          dv = hs.GetDeviceByRef(dv_ref)
        End If

        Call WriteMessage(String.Format("Creating new HomeSeer {0} device.", dv_name), MessageType.Debug)

      End If

      '
      ' Store the UUID for the device
      '
      Dim pdata As New clsPlugExtraData
      pdata.AddNamed("UUID", DeviceId)
      dv.PlugExtraData_Set(hs) = pdata

      '
      ' Define the HomeSeer device
      '
      dv.Address(hs) = dv_addr
      dv.Interface(hs) = IFACE_NAME
      dv.InterfaceInstance(hs) = Instance

      '
      ' Update location properties on new devices only
      '
      If bDeviceExists = False Then
        dv.Location(hs) = IFACE_NAME & " Plugin"
        dv.Location2(hs) = String.Format("Russound{0} {1}", device_id, zone_name)
      End If

      '
      ' The following simply shows up in the device properties but has no other use
      '
      dv.Device_Type_String(hs) = dv_type

      '
      ' Set the DeviceTypeInfo
      '
      Dim DT As New DeviceTypeInfo
      DT.Device_API = DeviceTypeInfo.eDeviceAPI.Plug_In
      DT.Device_Type = DeviceTypeInfo.eDeviceType_Plugin.Root
      dv.DeviceType_Set(hs) = DT

      '
      ' Make this device a child of the root
      '
      If dv.Relationship(hs) <> Enums.eRelationship.Child Then

        dv.AssociatedDevice_ClearAll(hs)
        Dim dvp_ref As Integer = CreateAudioRootDevice(DeviceId, device_desc, zone_name, dv_ref)
        If dvp_ref > 0 Then
          dv.AssociatedDevice_Add(hs, dvp_ref)
        End If
        dv.Relationship(hs) = Enums.eRelationship.Child

      End If

      '
      ' Update the last change date
      ' 
      dv.Last_Change(hs) = DateTime.Now

      DevicePairs.Clear()
      Select Case device_type
        '
        ' Master Controls
        '
        Case "audiozone-power"
          '
          ' Device can be controlled
          '
          'DevicePairs.Add(New hspi_device_pairs(-1, "Toggle", "audio_zone_on.png", HomeSeerAPI.ePairStatusControl.Control))
          DevicePairs.Add(New hspi_device_pairs(0, "Off", "audio_zone_off.png", HomeSeerAPI.ePairStatusControl.Both))
          DevicePairs.Add(New hspi_device_pairs(1, "On", "audio_zone_off.png", HomeSeerAPI.ePairStatusControl.Both))

        Case "audiozone-source"
          '
          ' Device can be controlled
          '
          DeviceShowValues = True

          '
          ' Define the default image
          '
          Dim strImage As String = "audio_zone_source.png"

          '
          ' Add input values
          '
          For input As Integer = 0 To 5
            Dim strValue As String = GetZoneSourceName(input)
            DevicePairs.Add(New hspi_device_pairs(input, strValue, strImage, HomeSeerAPI.ePairStatusControl.Both))
          Next

        Case "audiozone-source-controls"

          '
          ' Device can be controlled
          '
          DeviceShowValues = True

          '
          ' Define the default image
          '
          Dim strImage As String = "audio_zone_source_control.png"

          DevicePairs.Add(New hspi_device_pairs(0, "Controls", strImage, HomeSeerAPI.ePairStatusControl.Both))

          Dim itemValues As Array = System.Enum.GetValues(GetType(Keycodes))
          Dim itemNames As Array = System.Enum.GetNames(GetType(Keycodes))

          For i As Integer = 0 To itemNames.Length - 1
            Dim value As String = itemValues(i)
            Dim name As String = itemNames(i)
            DevicePairs.Add(New hspi_device_pairs(value, name, strImage, HomeSeerAPI.ePairStatusControl.Both))
          Next

        Case "audiozone-partymode"
          '
          ' Device can be controlled
          '
          DeviceShowValues = True

          '
          ' Define the default image
          '
          Dim strImage As String = "audio_zone_partymode.png"

          '
          ' Add input values
          '
          DevicePairs.Add(New hspi_device_pairs(-1, "Toggle", "audio_zone_party_mode.png", HomeSeerAPI.ePairStatusControl.Control))
          DevicePairs.Add(New hspi_device_pairs(0, "Off", "audio_zone_party_mode.png", HomeSeerAPI.ePairStatusControl.Both))
          DevicePairs.Add(New hspi_device_pairs(1, "On", "audio_zone_party_mode.png", HomeSeerAPI.ePairStatusControl.Both))
          DevicePairs.Add(New hspi_device_pairs(2, "Master", "audio_zone_party_mode.png", HomeSeerAPI.ePairStatusControl.Both))

        Case "audiozone-dnd"
          '
          ' Device can be controlled
          '
          DeviceShowValues = True

          '
          ' Define the default image
          '
          Dim strImage As String = "audio_zone_dnd.png"

          '
          ' Add input values
          '
          DevicePairs.Add(New hspi_device_pairs(-1, "Toggle", "audio_zone_dnd.png", HomeSeerAPI.ePairStatusControl.Control))
          DevicePairs.Add(New hspi_device_pairs(0, "Off", "audio_zone_dnd.png", HomeSeerAPI.ePairStatusControl.Both))
          DevicePairs.Add(New hspi_device_pairs(1, "On", "audio_zone_dnd.png", HomeSeerAPI.ePairStatusControl.Both))

        Case "audiozone-turnon-vol"
          '
          ' Device can be controlled
          '
          DeviceShowValues = True

          '
          ' Define the default image
          '
          Dim strImage As String = "audio_zone_volume.png"

          DevicePairs.Add(New hspi_device_pairs(-2, "Decrease", strImage, HomeSeerAPI.ePairStatusControl.Control))
          DevicePairs.Add(New hspi_device_pairs(-1, "Increase", strImage, HomeSeerAPI.ePairStatusControl.Control))

          For i As Integer = &H00 To &H32
            Dim strVolume As String = GetZoneVolumeLevel(i)
            DevicePairs.Add(New hspi_device_pairs(i, strVolume, strImage, HomeSeerAPI.ePairStatusControl.Both))
          Next

        '
        ' Zone Audio Controls
        '
        Case "audiozone-volume"
          '
          ' Device can be controlled
          '
          DeviceShowValues = True

          '
          ' Define the default image
          '
          Dim strImage As String = "audio_zone_volume.png"

          For i As Integer = &H00 To &H32
            Dim strVolume As String = GetZoneVolumeLevel(i)
            DevicePairs.Add(New hspi_device_pairs(i, strVolume, strImage, HomeSeerAPI.ePairStatusControl.Both))
          Next

        Case "audiozone-bass"
          '
          ' Device can be controlled
          '
          DeviceShowValues = True

          '
          ' Define the default image
          '
          Dim strImage As String = "audio_zone_bass.png"

          DevicePairs.Add(New hspi_device_pairs(-2, "Decrease", strImage, HomeSeerAPI.ePairStatusControl.Control))
          DevicePairs.Add(New hspi_device_pairs(-1, "Increase", strImage, HomeSeerAPI.ePairStatusControl.Control))

          For i As Integer = &H00 To &H14
            Dim strVolume As String = GetZoneBassLevel(i)
            DevicePairs.Add(New hspi_device_pairs(i, strVolume, strImage, HomeSeerAPI.ePairStatusControl.Both))
          Next

        Case "audiozone-treble"
          '
          ' Device can be controlled
          '
          DeviceShowValues = True

          '
          ' Define the default image
          '
          Dim strImage As String = "audio_zone_treble.png"

          DevicePairs.Add(New hspi_device_pairs(-2, "Decrease", strImage, HomeSeerAPI.ePairStatusControl.Control))
          DevicePairs.Add(New hspi_device_pairs(-1, "Increase", strImage, HomeSeerAPI.ePairStatusControl.Control))

          For i As Integer = &H00 To &H14
            Dim strVolume As String = GetZoneTrebleLevel(i)
            DevicePairs.Add(New hspi_device_pairs(i, strVolume, strImage, HomeSeerAPI.ePairStatusControl.Both))
          Next

        Case "audiozone-balance"
          '
          ' Device can be controlled
          '
          DeviceShowValues = True

          '
          ' Define the default image
          '
          Dim strImage As String = "audio_zone_balance.png"

          DevicePairs.Add(New hspi_device_pairs(-2, "More Left", strImage, HomeSeerAPI.ePairStatusControl.Control))
          DevicePairs.Add(New hspi_device_pairs(-1, "More Right", strImage, HomeSeerAPI.ePairStatusControl.Control))

          For i As Integer = &H00 To &H14
            Dim strVolume As String = GetZoneBalanceLevel(i)
            DevicePairs.Add(New hspi_device_pairs(i, strVolume, strImage, HomeSeerAPI.ePairStatusControl.Both))
          Next

        Case "audiozone-loudness"
          '
          ' Device can be controlled
          '
          DeviceShowValues = True

          '
          ' Define the default image
          '
          Dim strImage As String = "audio_zone_loudness.png"

          '
          ' Add input values
          '
          DevicePairs.Add(New hspi_device_pairs(-1, "Toggle", "audio_zone_loudness.png", HomeSeerAPI.ePairStatusControl.Control))
          DevicePairs.Add(New hspi_device_pairs(0, "Off", "audio_zone_loudness.png", HomeSeerAPI.ePairStatusControl.Both))
          DevicePairs.Add(New hspi_device_pairs(1, "On", "audio_zone_loudness.png", HomeSeerAPI.ePairStatusControl.Both))

        '
        ' Zone Keypads
        '
        Case "audiozone-keypad-controls"
          '
          ' Device can be controlled
          '
          DeviceShowValues = True

          '
          ' Define the default image
          '
          Dim strImage As String = "audio_zone_keypad_controls.png"

          DevicePairs.Add(New hspi_device_pairs(0, "Controls", strImage, HomeSeerAPI.ePairStatusControl.Both))

          Dim itemValues As Array = System.Enum.GetValues(GetType(KeypadEvents))
          Dim itemNames As Array = System.Enum.GetNames(GetType(KeypadEvents))

          For i As Integer = 0 To itemNames.Length - 1
            Dim value As String = itemValues(i)
            Dim name As String = itemNames(i)
            DevicePairs.Add(New hspi_device_pairs(value, name, strImage, HomeSeerAPI.ePairStatusControl.Both))
          Next

        Case "audiozone-keypad-bg-color"
          '
          ' Device can be controlled
          '
          DeviceShowValues = True

          '
          ' Define the default image
          '
          Dim strImage As String = "audio_zone_bg_color.png"

          '
          ' Add input values
          '
          DevicePairs.Add(New hspi_device_pairs(-1, "Toggle", "audio_zone_bg_color.png", HomeSeerAPI.ePairStatusControl.Control))
          DevicePairs.Add(New hspi_device_pairs(0, "Off", "audio_zone_bg_color.png", HomeSeerAPI.ePairStatusControl.Both))
          DevicePairs.Add(New hspi_device_pairs(1, "Amber", "audio_zone_bg_color.png", HomeSeerAPI.ePairStatusControl.Both))
          DevicePairs.Add(New hspi_device_pairs(2, "Green", "audio_zone_bg_color.png", HomeSeerAPI.ePairStatusControl.Both))

        '
        ' Zone Misc
        '
        Case "audiozone-system-onstate"
          '
          ' Device can be controlled
          '
          DeviceShowValues = False

          '
          ' Define the default image
          '
          Dim strImage As String = "audio_zone_on_state.png"

          '
          ' Add input values
          '
          DevicePairs.Add(New hspi_device_pairs(0, "All Zones Off", "audio_zone_off_state.png", HomeSeerAPI.ePairStatusControl.Status))
          DevicePairs.Add(New hspi_device_pairs(1, "Any Zone is On", "audio_zone_on_state.png", HomeSeerAPI.ePairStatusControl.Both))

        Case "audiozone-source-shared"
          '
          ' Device can be controlled
          '
          DeviceShowValues = False

          '
          ' Define the default image
          '
          Dim strImage As String = "audio_zone_source_shared.png"

          '
          ' Add input values
          '
          DevicePairs.Add(New hspi_device_pairs(0, "Not Shared", "audio_zone_source_shared.png", HomeSeerAPI.ePairStatusControl.Status))
          DevicePairs.Add(New hspi_device_pairs(1, "Shared", "audio_zone_source_shared.png", HomeSeerAPI.ePairStatusControl.Both))

        Case "audiozone-keypad-display"
          '
          ' Device can be controlled
          '
          DeviceShowValues = False

          '
          ' Define the default image
          '
          Dim strImage As String = "audio_zone_keypad_message.png"

          '
          ' Add input values
          '
          DevicePairs.Add(New hspi_device_pairs(0, "Display", "audio_zone_keypad_message.png", HomeSeerAPI.ePairStatusControl.Status))

      End Select

      '
      ' Add the Graphic Value Pairs
      '
      If DevicePairs.Count > 0 Then

        For Each Pair As hspi_device_pairs In DevicePairs

          Dim VSPair As VSPair = New VSPair(Pair.Type)
          VSPair.PairType = VSVGPairType.SingleValue
          VSPair.Value = Pair.Value
          VSPair.Status = Pair.Status
          VSPair.Render = Enums.CAPIControlType.Values
          hs.DeviceVSP_AddPair(dv_ref, VSPair)

          Dim VGPair As VGPair = New VGPair()
          VGPair.PairType = VSVGPairType.SingleValue
          VGPair.Set_Value = Pair.Value
          VGPair.Graphic = String.Format("{0}{1}", gImageDir, Pair.Image)
          hs.DeviceVGP_AddPair(dv_ref, VGPair)

        Next

      End If

      '
      ' Update the Device Misc Bits
      '
      If DeviceShowValues = True Then
        dv.MISC_Set(hs, Enums.dvMISC.SHOW_VALUES)
      End If

      If bDeviceExists = False Then
        dv.MISC_Set(hs, Enums.dvMISC.NO_LOG)
      End If

      dv.Status_Support(hs) = False

      hs.SaveEventsDevices()

      '
      ' Update the Audio Device
      '
      If bUpdateOnly = False Then

        If gAudioDevices.Any(Function(s) s.DeviceId = device_id) = True Then
          Dim AudioDevice As hspi_audio_device = gAudioDevices.Find(Function(s) s.DeviceId = device_id)

          Dim AudioZone As hspi_audio_zone = AudioDevice.AudioZones.Find(Function(s) s.DeviceAddr = dv_addr)
          If Not IsNothing(AudioZone) Then
            AudioZone.DeviceExists = True
          End If

          SetDeviceValue(dv_addr, AudioZone.GetPropertyValue(device_type))
        End If

      End If

      Return ""

    Catch pEx As Exception
      '
      ' Process program exception
      '
      Call ProcessError(pEx, "CreateAudioZoneDevice()")
      Return "Failed to create HomeSeer device due to error."
    End Try

  End Function

  ''' <summary>
  ''' 
  ''' </summary>
  ''' <param name="dv_addr"></param>
  ''' <param name="bUpdateOnly"></param>
  ''' <returns></returns>
  Public Function CreateTunerDevice(ByVal dv_addr As String, ByVal bUpdateOnly As Boolean)

    Dim dv As Scheduler.Classes.DeviceClass

    Dim dv_ref As Integer = 0
    Dim dv_misc As Integer = 0

    Dim dv_name As String = String.Empty
    Dim dv_type As String = String.Empty

    Dim DeviceShowValues As Boolean = True
    Dim DevicePairs As New ArrayList

    Try

      '
      ' Set the local variables
      '
      Dim regexPattern As String = "Russound(?<dd>(\d+))\.(?<cc>(\d+))\.(?<tt>(\d+))-(?<type>(.+))"

      Dim device_id As String = Regex.Match(dv_addr, regexPattern).Groups("dd").ToString()
      Dim cc As String = Regex.Match(dv_addr, regexPattern).Groups("cc").ToString()
      Dim tt As String = Regex.Match(dv_addr, regexPattern).Groups("tt").ToString()
      Dim device_type As String = Regex.Match(dv_addr, regexPattern).Groups("type").ToString()

      Dim tuner_id As String = String.Format("{0}.{1}.{2}", device_id, cc, tt)
      Dim tuner_number As Integer = tt + 1
      Dim tuner_name As String = String.Format("Tuner{0}", tuner_number.ToString.PadLeft(2, "0"))
      Dim device_desc As String = GetKeyFriendlyName(device_type)

      dv_type = String.Format("Russound Tuner {0}", device_desc)
      dv_name = String.Format("Russound{0} {1} {2}", device_id, tuner_name, device_desc)

      dv = LocateDeviceByAddr(dv_addr)

      Dim bDeviceExists As Boolean = Not dv Is Nothing
      If bUpdateOnly = True And bDeviceExists = False Then Return ""

      If bDeviceExists = True Then
        '
        ' Lets use the existing device
        '
        dv_addr = dv.Address(hs)
        dv_ref = dv.Ref(hs)

        Call WriteMessage(String.Format("Updating existing HomeSeer {0} device.", dv_name), MessageType.Debug)

      Else
        '
        ' Create A HomeSeer Device
        '
        dv_ref = hs.NewDeviceRef(dv_name)
        If dv_ref > 0 Then
          dv = hs.GetDeviceByRef(dv_ref)
        End If

        Call WriteMessage(String.Format("Creating new HomeSeer {0} device.", dv_name), MessageType.Debug)

      End If

      '
      ' Store the UUID for the device
      '
      Dim pdata As New clsPlugExtraData
      pdata.AddNamed("UUID", DeviceId)
      dv.PlugExtraData_Set(hs) = pdata

      '
      ' Define the HomeSeer device
      '
      dv.Address(hs) = dv_addr
      dv.Interface(hs) = IFACE_NAME
      dv.InterfaceInstance(hs) = Instance

      '
      ' Update location properties on new devices only
      '
      If bDeviceExists = False Then
        dv.Location(hs) = IFACE_NAME & " Plugin"
        dv.Location2(hs) = String.Format("Russound{0} {1}", device_id, tuner_name)
      End If

      '
      ' The following simply shows up in the device properties but has no other use
      '
      dv.Device_Type_String(hs) = dv_type

      '
      ' Set the DeviceTypeInfo
      '
      Dim DT As New DeviceTypeInfo
      DT.Device_API = DeviceTypeInfo.eDeviceAPI.Plug_In
      DT.Device_Type = DeviceTypeInfo.eDeviceType_Plugin.Root
      dv.DeviceType_Set(hs) = DT

      '
      ' Make this device a child of the root
      '
      If dv.Relationship(hs) <> Enums.eRelationship.Child Then

        dv.AssociatedDevice_ClearAll(hs)
        Dim dvp_ref As Integer = CreateAudioRootDevice(DeviceId, device_desc, tuner_name, dv_ref)
        If dvp_ref > 0 Then
          dv.AssociatedDevice_Add(hs, dvp_ref)
        End If
        dv.Relationship(hs) = Enums.eRelationship.Child

      End If

      '
      ' Update the last change date
      ' 
      dv.Last_Change(hs) = DateTime.Now

      DevicePairs.Clear()
      Select Case device_type

        Case "tuner-radio-controls"
          '
          ' Device can be controlled
          '
          DeviceShowValues = True

          '
          ' Define the default image
          '
          Dim strImage As String = "tuner_radio_controls.png"

          DevicePairs.Add(New hspi_device_pairs(0, "Controls", strImage, HomeSeerAPI.ePairStatusControl.Both))

          Dim itemValues As Array = System.Enum.GetValues(GetType(TunerCommand))
          Dim itemNames As Array = System.Enum.GetNames(GetType(TunerCommand))

          For i As Integer = 0 To itemNames.Length - 1
            Dim value As String = itemValues(i)
            Select Case value
              Case &H0A, &H01 - &H09, &H1B, &H46, &H49, &H4A, &H13, &H33, &H34, &H42, &H43
                Dim name As String = itemNames(i)
                DevicePairs.Add(New hspi_device_pairs(value, name, strImage, HomeSeerAPI.ePairStatusControl.Control))
            End Select

          Next

        Case "tuner-power"
          '
          ' Device can be controlled
          '
          DeviceShowValues = True

          '
          ' Define the default image
          '
          Dim strImage As String = "tuner_power.png"

          DevicePairs.Add(New hspi_device_pairs(-2, "Off", strImage, HomeSeerAPI.ePairStatusControl.Control))
          DevicePairs.Add(New hspi_device_pairs(-1, "On", strImage, HomeSeerAPI.ePairStatusControl.Control))
          DevicePairs.Add(New hspi_device_pairs(0, "Power", strImage, HomeSeerAPI.ePairStatusControl.Both))

        Case "tuner-preset"
          '
          ' Device can be controlled
          '
          DeviceShowValues = True

          '
          ' Define the default image
          '
          Dim strImage As String = "tuner_preset.png"

          DevicePairs.Add(New hspi_device_pairs(-2, "Preset Down", strImage, HomeSeerAPI.ePairStatusControl.Control))
          DevicePairs.Add(New hspi_device_pairs(-1, "Preset Up", strImage, HomeSeerAPI.ePairStatusControl.Control))
          DevicePairs.Add(New hspi_device_pairs(0, "Preset", strImage, HomeSeerAPI.ePairStatusControl.Both))
          For i As Byte = 1 To 6
            Dim strPrefix As String = String.Format("Preset {0}", i.ToString)
            DevicePairs.Add(New hspi_device_pairs(i, strPrefix, strImage, HomeSeerAPI.ePairStatusControl.Control))
          Next

        Case "tuner-bank"
          '
          ' Device can be controlled
          '
          DeviceShowValues = True

          '
          ' Define the default image
          '
          Dim strImage As String = "tuner_bank.png"

          DevicePairs.Add(New hspi_device_pairs(-2, "Bank Down", strImage, HomeSeerAPI.ePairStatusControl.Control))
          DevicePairs.Add(New hspi_device_pairs(-1, "Bank Up", strImage, HomeSeerAPI.ePairStatusControl.Control))
          DevicePairs.Add(New hspi_device_pairs(0, "Bank", strImage, HomeSeerAPI.ePairStatusControl.Both))
          For i As Byte = 1 To 6
            Dim strPrefix As String = String.Format("Bank {0}", i.ToString)
            DevicePairs.Add(New hspi_device_pairs(i, strPrefix, strImage, HomeSeerAPI.ePairStatusControl.Control))
          Next

        Case "tuner-radio-frequency"
          '
          ' Device can be controlled
          '
          DeviceShowValues = True

          '
          ' Define the default image
          '
          Dim strImage As String = "tuner_frequency.png"

          DevicePairs.Add(New hspi_device_pairs(-2, "Frequency Down", strImage, HomeSeerAPI.ePairStatusControl.Control))
          DevicePairs.Add(New hspi_device_pairs(-1, "Frequency Up", strImage, HomeSeerAPI.ePairStatusControl.Control))
          DevicePairs.Add(New hspi_device_pairs(0, "Frequency", strImage, HomeSeerAPI.ePairStatusControl.Both))

          For i As Integer = 88 To 108 Step 1
            For j As Double = 0 To 1 Step 0.1
              Dim value As Double = i + j
              Dim strValue As String = value.ToString("000.0")
              Dim strKey As String = strValue.ToString.Replace(".", "")
              Select Case strKey
                Case "1090"
                Case Else
                  DevicePairs.Add(New hspi_device_pairs(strKey, strValue, strImage, HomeSeerAPI.ePairStatusControl.Control))
              End Select

            Next

          Next

        Case "tuner-xm-channel"
          '
          ' Device can be controlled
          '
          DeviceShowValues = True

          '
          ' Define the default image
          '
          Dim strImage As String = "tuner_xm_channel.png"

          DevicePairs.Add(New hspi_device_pairs(-2, "Channel Down", strImage, HomeSeerAPI.ePairStatusControl.Control))
          DevicePairs.Add(New hspi_device_pairs(-1, "Channel Up", strImage, HomeSeerAPI.ePairStatusControl.Control))
          DevicePairs.Add(New hspi_device_pairs(0, "XM Channel", strImage, HomeSeerAPI.ePairStatusControl.Both))

          For channel As Integer = 1 To 236
            Dim strChannelName As String = GetXMChannelName(channel)
            If strChannelName.Length > 0 Then
              DevicePairs.Add(New hspi_device_pairs(channel, strChannelName, strImage, HomeSeerAPI.ePairStatusControl.Control))
            End If
          Next

        Case "tuner-xm-category"
          '
          ' Device can be controlled
          '
          DeviceShowValues = True

          '
          ' Define the default image
          '
          Dim strImage As String = "tuner_xm_category.png"

          DevicePairs.Add(New hspi_device_pairs(-2, "Category Down", strImage, HomeSeerAPI.ePairStatusControl.Control))
          DevicePairs.Add(New hspi_device_pairs(-1, "Category Up", strImage, HomeSeerAPI.ePairStatusControl.Control))
          DevicePairs.Add(New hspi_device_pairs(0, "Category", strImage, HomeSeerAPI.ePairStatusControl.Both))

          'Dim itemValues As Array = System.Enum.GetValues(GetType(XMCategories))
          'Dim itemNames As Array = System.Enum.GetNames(GetType(XMCategories))

          'For i As Integer = 0 To itemNames.Length - 1
          '  Dim value As String = itemValues(i)
          '  Dim name As String = itemNames(i)
          '  If value <> &H00 Then
          '    DevicePairs.Add(New hspi_device_pairs(value, name, strImage, HomeSeerAPI.ePairStatusControl.Control))
          '  End If

          'Next

        Case "tuner-xm-category-channel"
          '
          ' Device can be controlled
          '
          DeviceShowValues = True

          '
          ' Define the default image
          '
          Dim strImage As String = "tuner_xm_channel.png"

          DevicePairs.Add(New hspi_device_pairs(-2, "Channel Down", strImage, HomeSeerAPI.ePairStatusControl.Control))
          DevicePairs.Add(New hspi_device_pairs(-1, "Channel Up", strImage, HomeSeerAPI.ePairStatusControl.Control))
          DevicePairs.Add(New hspi_device_pairs(0, "Category Channel", strImage, HomeSeerAPI.ePairStatusControl.Both))

      End Select

      '
      ' Add the Graphic Value Pairs
      '
      If DevicePairs.Count > 0 Then

        For Each Pair As hspi_device_pairs In DevicePairs

          Dim VSPair As VSPair = New VSPair(Pair.Type)
          VSPair.PairType = VSVGPairType.SingleValue
          VSPair.Value = Pair.Value
          VSPair.Status = Pair.Status
          VSPair.Render = Enums.CAPIControlType.Values
          hs.DeviceVSP_AddPair(dv_ref, VSPair)

          Dim VGPair As VGPair = New VGPair()
          VGPair.PairType = VSVGPairType.SingleValue
          VGPair.Set_Value = Pair.Value
          VGPair.Graphic = String.Format("{0}{1}", gImageDir, Pair.Image)
          hs.DeviceVGP_AddPair(dv_ref, VGPair)

        Next

      End If

      '
      ' Update the Device Misc Bits
      '
      If DeviceShowValues = True Then
        dv.MISC_Set(hs, Enums.dvMISC.SHOW_VALUES)
      End If

      If bDeviceExists = False Then
        dv.MISC_Set(hs, Enums.dvMISC.NO_LOG)
      End If

      dv.Status_Support(hs) = False

      hs.SaveEventsDevices()

      Return ""

    Catch pEx As Exception
      '
      ' Process program exception
      '
      Call ProcessError(pEx, "CreateAudioZoneDevice()")
      Return "Failed to create HomeSeer device due to error."
    End Try

  End Function

  ''' <summary>
  ''' Returns the XM Channel Name
  ''' </summary>
  ''' <param name="channel"></param>
  ''' <returns></returns>
  Function GetXMChannelName(channel As Integer) As String

    Select Case channel
      Case 1 : Return "1 - Preview"
      Case 2 : Return "2 - Sirius XM Hits 1"
      Case 3 : Return "3 - Venus"
      Case 4 : Return "4 - Pitbull's Globalization"
      Case 5 : Return "5 - 50s on 5"
      Case 6 : Return "6 - 60s on 6"
      Case 7 : Return "7 - 70s on 7"
      Case 8 : Return "8 - 80s on 8"
      Case 9 : Return "9 - 90s on 9"
      Case 10 : Return "10 - Pop 2K"
      Case 11 : Return "11 - KIIS-FM"
      Case 12 : Return "12 - Z-100 FM"
      Case 13 : Return "13 - Velvet"
      Case 14 : Return "14 - The Coffee House"
      Case 15 : Return "15 - The Pulse"
      Case 16 : Return "16 - The Blend"
      Case 17 : Return "17 - SiriusXM Love"
      Case 18 : Return "18 - SiriusXM Limited Edition"
      Case 19 : Return "19 - Elvis Radio"
      Case 20 : Return "20 - E Street Radio"
      Case 21 : Return "21 - Underground Garage"
      Case 22 : Return "22 - Pearl Jam Radio"
      Case 23 : Return "23 - Grateful Dead Channel"
      Case 24 : Return "24 - Radio Margaritaville"
      Case 25 : Return "25 - Classic Rewind"
      Case 26 : Return "26 - Classic Vinyl"
      Case 27 : Return "27 - Deep Tracks"
      Case 28 : Return "28 - The Spectrum"
      Case 29 : Return "29 - Jam On"
      Case 30 : Return "30 - The Loft"
      Case 31 : Return "31 - Tom Petty Radio"
      Case 32 : Return "32 - The Bridge"
      Case 33 : Return "33 - 1st Wave"
      Case 34 : Return "34 - Lithium"
      Case 35 : Return "35 - Sirius XMU"
      Case 36 : Return "36 - Alt Nation"
      Case 37 : Return "37 - Octane"
      Case 38 : Return "38 - Boneyard"
      Case 39 : Return "39 - Hair Nation"
      Case 40 : Return "40 - Liquid Metal"
      Case 41 : Return "41 - Faction"
      Case 42 : Return "42 - The Joint"
      Case 43 : Return "43 - Backspin"
      Case 44 : Return "44 - HipHop Nation"
      Case 45 : Return "45 - Shade 45"
      Case 46 : Return "46 - The Heat"
      Case 47 : Return "47 - SiriusXM Fly"
      Case 48 : Return "48 - Heart and Soul"
      Case 49 : Return "49 - Soul Town"
      Case 50 : Return "50 - The Groove"
      Case 51 : Return "51 - BPM"
      Case 52 : Return "52 - Electric Area"
      Case 53 : Return "53 - Sirius XM Chill"
      Case 54 : Return "54 - Studio 54 Radio"
      Case 55 : Return "55 - The Garth Channel"
      Case 56 : Return "56 - The Highway"
      Case 57 : Return "57 - No Shoes Radio"
      Case 58 : Return "58 - Prime Country"
      Case 59 : Return "59 - Willies Roadhouse"
      Case 60 : Return "60 - Outlaw Country"
      Case 61 : Return "61 - Y2Kountry"
      Case 62 : Return "62 - Bluegrass Junction"
      Case 63 : Return "63 - The Message"
      Case 64 : Return "64 - Praise"
      Case 65 : Return "65 - enLighten"
      Case 66 : Return "66 - Watercolors"
      Case 67 : Return "67 - Real Jazz"
      Case 68 : Return "68 - Spa"
      Case 69 : Return "69 - Escape"
      Case 70 : Return "70 - BB Kings Bluesville"
      Case 71 : Return "71 - Siriusly Sinatra"
      Case 72 : Return "72 - On Broadway"
      Case 73 : Return "73 - 40's Junction"
      Case 74 : Return "74 - Met Opera Radio"
      Case 76 : Return "76 - Symphony Hall"
      Case 77 : Return "77 - KIDZ BOP Radio"
      Case 78 : Return "78 - Kids Place Live"
      Case 79 : Return "79 - Radio Disney"
      Case 80 : Return "80 - ESPN Radio"
      Case 81 : Return "81 - ESPN Xtra"
      Case 82 : Return "82 - Mad Dog Radio"
      Case 83 : Return "83 - Bleacher Report Radio"
      Case 84 : Return "84 - College Sports Nation"
      Case 85 : Return "85 - SiriusXM FC"
      Case 86 : Return "86 - SiriusXM NBA Radio"
      Case 87 : Return "87 - Fantasy Sports Radio"
      Case 88 : Return "88 - NFL Radio"
      Case 89 : Return "89 - MLB Network Radio"
      Case 90 : Return "90 - NASCAR Radio"
      Case 91 : Return "91 - NHL Network Radio"
      Case 92 : Return "92 - PGA Tour Network"
      Case 93 : Return "93 - SiriusXM Rush"
      Case 94 : Return "94 - SiriusXM Comedy Greats"
      Case 95 : Return "95 - Comedy Central Radio"
      Case 96 : Return "96 - The Foxxhole"
      Case 97 : Return "97 - Jeff and Larry's Comedy Roundup"
      Case 98 : Return "98 - Laugh USA"
      Case 99 : Return "99 - Raw Dog Comedy"
      Case 100 : Return "100 - Howard Stern 100"
      Case 101 : Return "101 - Howard Stern 101"
      Case 102 : Return "102 - Radio Andy"
      Case 103 : Return "103 - Opie and Anthony"
      Case 105 : Return "105 - Entertainment Weekly"
      Case 106 : Return "106 - SiriusXM 106"
      Case 108 : Return "108 - Today Show Radio"
      Case 109 : Return "109 - Sirius XM Stars"
      Case 110 : Return "110 - Doctor Radio"
      Case 111 : Return "111 - Wharton Business Radio"
      Case 112 : Return "112 - CNBC"
      Case 113 : Return "113 - FOX Business"
      Case 114 : Return "114 - FOX News"
      Case 115 : Return "115 - FOX News Headlines 24/7"
      Case 116 : Return "116 - CNN"
      Case 117 : Return "117 - HLN"
      Case 118 : Return "118 - MSNBC"
      Case 119 : Return "119 - Bloomberg Radio"
      Case 120 : Return "120 - BBC World Service"
      Case 121 : Return "121 - SiriusXM Insight"
      Case 122 : Return "122 - NPR Now"
      Case 123 : Return "123 - PRX"
      Case 124 : Return "124 - POTUS"
      Case 125 : Return "125 - Sirius XM Patriot"
      Case 126 : Return "126 - Urban View"
      Case 127 : Return "127 - Sirius XM Progress"
      Case 128 : Return "128 - Joel Osteen Radio"
      Case 129 : Return "129 - Catholic Channel"
      Case 131 : Return "131 - Family Talk"
      Case 132 : Return "132 - Traffic: BOS/PHI/PIT/ATL"
      Case 133 : Return "133 - Traffic: NY"
      Case 134 : Return "134 - Traffic: DC/BAL/MIA"
      Case 135 : Return "135 - Traffic: CHI/DET"
      Case 136 : Return "136 - Traffic: DFW/STL/MSP"
      Case 137 : Return "137 - Traffic: HOU/PHX/TB/SF/SEA"
      Case 138 : Return "138 - Traffic: LSV/LA/SDG"
      Case 141 : Return "141 - HUR Voices"
      Case 142 : Return "142 - HBCU Radio"
      Case 143 : Return "143 - BYU Radio"
      Case 144 : Return "144 - Korea Today"
      Case 146 : Return "146 - Road Dog Trucking"
      Case 147 : Return "147 - Rural Radio"
      Case 148 : Return "148 - Radio Classics"
      Case 152 : Return "152 - En Vivo"
      Case 153 : Return "153 - Cristina Radio"
      Case 154 : Return "154 - Inspirate"
      Case 155 : Return "155 - CNN en Espanol"
      Case 157 : Return "157 - ESPN Deportes Radio"
      Case 158 : Return "158 - Caliente"
      Case 159 : Return "159 - MLB en Espanol"
      Case 162 : Return "162 - CBC Radio 3"
      Case 163 : Return "163 - Chansons"
      Case 166 : Return "166 - Franco Country"
      Case 167 : Return "167 - Canada Talks"
      Case 168 : Return "168 - Canada Laughs"
      Case 169 : Return "169 - CBC Radio One"
      Case 170 : Return "170 - CBC Premiere"
      Case 171 : Return "171 - CBC Country"
      Case 172 : Return "172 - Canada 360"
      Case 173 : Return "173 - The Verge"
      Case 174 : Return "174 - Influence Franco"
      Case 176 : Return "176 - MLB Play by Play 01"
      Case 177 : Return "177 - MLB Play by Play 02"
      Case 178 : Return "178 - MLB Play by Play 03"
      Case 179 : Return "179 - MLB Play by Play 04"
      Case 180 : Return "180 - MLB Play by Play 05"
      Case 181 : Return "181 - MLB Play by Play 06"
      Case 182 : Return "182 - MLB Play by Play 07"
      Case 183 : Return "183 - MLB Play by Play 08"
      Case 184 : Return "184 - MLB Play by Play 09"
      Case 185 : Return "185 - MLB Play by Play 10"
      Case 186 : Return "186 - MLB Play by Play 11"
      Case 187 : Return "187 - MLB Play by Play 12"
      Case 188 : Return "188 - MLB Play by Play 13"
      Case 189 : Return "189 - MLB Play by Play 14"
      Case 190 : Return "190 - NCAA Play by Play 01"
      Case 191 : Return "191 - NCAA Play by Play 02"
      Case 192 : Return "192 - NCAA Play by Play 03"
      Case 193 : Return "193 - NCAA Play by Play 04"
      Case 194 : Return "194 - NCAA Play by Play 05"
      Case 195 : Return "195 - NCAA Play by Play 06"
      Case 196 : Return "196 - NCAA Play by Play 07"
      Case 197 : Return "197 - NCAA Play by Play 08"
      Case 198 : Return "198 - NCAA Play by Play 09"
      Case 199 : Return "199 - NCAA Play by Play 10"
      Case 200 : Return "200 - NCAA Play by Play 11"
      Case 201 : Return "201 - Sports Play by Play 01"
      Case 202 : Return "202 - Sports Play by Play 02"
      Case 203 : Return "203 - Sports Play by Play 03"
      Case 204 : Return "204 - Sports Play by Play 04"
      Case 205 : Return "205 - Sports Play by Play 05"
      Case 206 : Return "206 - Sports Play by Play 06"
      Case 207 : Return "207 - Sports Play by Play 07"
      Case 209 : Return "209 - Verizon IndyCar"
      Case 210 : Return "210 - Sports Play by Play 09"
      Case 211 : Return "211 - Sports Play by Play 10"
      Case 212 : Return "212 - NBA Play by Play 01"
      Case 213 : Return "213 - NBA Play by Play 02"
      Case 214 : Return "214 - NBA Play by Play 03"
      Case 215 : Return "215 - NBA Play by Play 04"
      Case 216 : Return "216 - NBA Play by Play 05"
      Case 217 : Return "217 - NBA Play by Play 06"
      Case 219 : Return "219 - NHL Play by Play 01"
      Case 220 : Return "220 - NHL Play by Play 02"
      Case 221 : Return "221 - NHL Play by Play 03"
      Case 222 : Return "222 - NHL Play by Play 04"
      Case 223 : Return "223 - NHL Play by Play 05"
      Case 225 : Return "225 - NFL Play by Play 01"
      Case 226 : Return "226 - NFL Play by Play 02"
      Case 227 : Return "227 - NFL Play by Play 03"
      Case 228 : Return "228 - NFL Play by Play 04"
      Case 229 : Return "229 - NFL Play by Play 05"
      Case 230 : Return "230 - NFL Play by Play 06"
      Case 231 : Return "231 - NFL Play by Play 07"
      Case 232 : Return "232 - NFL Play by Play 08"
      Case 233 : Return "233 - NFL Play by Play 09"
      Case 234 : Return "234 - NFL Play by Play 10"
      Case 235 : Return "235 - Sports Play by Play 11"
      Case 236 : Return "236 - Sports Play by Play 12"
      Case Else : Return String.Empty
    End Select

  End Function

  ''' <summary>
  ''' Format the zone source name
  ''' </summary>
  ''' <param name="input"></param>
  ''' <returns></returns>
  Function GetZoneSourceName(ByVal input As Integer) As String

    Return String.Format("Source {0}", input + 1)

  End Function

  ''' <summary>
  ''' Formats the Zone Volume Level
  ''' </summary>
  ''' <param name="level"></param>
  ''' <returns></returns>
  Private Function GetZoneVolumeLevel(ByVal level As Byte) As String

    Return String.Format("Volume {0}", level * 2)

  End Function

  ''' <summary>
  ''' Formats the Zone Bass Level
  ''' </summary>
  ''' <param name="level"></param>
  ''' <returns></returns>
  Private Function GetZoneBassLevel(ByVal level As Byte) As String

    Select Case level - 10
      Case < 0 : Return String.Format("Bass {0}", level - 10)
      Case > 0 : Return String.Format("Bass +{0}", level - 10)
      Case Else : Return "Bass Flat"
    End Select

  End Function

  ''' <summary>
  ''' Formats the Zone Treble Level
  ''' </summary>
  ''' <param name="level"></param>
  ''' <returns></returns>
  Private Function GetZoneTrebleLevel(ByVal level As Byte) As String

    Select Case level - 10
      Case < 0 : Return String.Format("Treble {0}", level - 10)
      Case > 0 : Return String.Format("Treble +{0}", level - 10)
      Case Else : Return "Treble Flat"
    End Select

  End Function

  ''' <summary>
  ''' Formats the Zone Balance Level
  ''' </summary>
  ''' <param name="level"></param>
  ''' <returns></returns>
  Private Function GetZoneBalanceLevel(ByVal level As Byte) As String

    Select Case level - 10
      Case < 0 : Return String.Format("Left {0}", level - 10)
      Case > 0 : Return String.Format("Right +{0}", level - 10)
      Case Else : Return "Centered"
    End Select

  End Function

#End Region

#Region "Russound String Tables"

  Private Function GetSourceName(ByVal index As Integer) As String

    Select Case index
      Case 0 : Return "AUX 1"
      Case 1 : Return "AUX 2"
      Case 2 : Return "AUX"
      Case 3 : Return "BLUES"
      Case 4 : Return "CABLE 1"
      Case 5 : Return "CABLE 2"
      Case 6 : Return "CABLE 3"
      Case 7 : Return "CABLE"
      Case 8 : Return "CD CHANGER"
      Case 9 : Return "CD CHANGER 1"
      Case 10 : Return "CD CHANGER 2"
      Case 11 : Return "CD CHANGER 3"
      Case 12 : Return "CD PLAYER"
      Case 13 : Return "CD PLAYER 1"
      Case 14 : Return "CD PLAYER 2"
      Case 15 : Return "CD PLAYER 3"
      Case 16 : Return "CLASSICAL"
      Case 17 : Return "COMPUTER"
      Case 18 : Return "COUNTRY"
      Case 19 : Return "DANCE MUSIC"
      Case 20 : Return "DIGITAL CABLE"
      Case 21 : Return "DSS RECEIVER"
      Case 22 : Return "DSS 1"
      Case 23 : Return "DSS 2"
      Case 24 : Return "DSS 3"
      Case 25 : Return "DVD CHANGER"
      Case 26 : Return "DVD CHANGER 1"
      Case 27 : Return "DVD CHANGER 2"
      Case 28 : Return "DVD CHANGER 3"
      Case 29 : Return "DVD PLAYER"
      Case 30 : Return "DVD PLAYER 1"
      Case 31 : Return "DVD PLAYER 2"
      Case 32 : Return "DVD PLAYER 3"
      Case 33 : Return "FRONT DOOR"
      Case 34 : Return "INTERNET RADIO"
      Case 35 : Return "JAZZ"
      Case 36 : Return "LASER DISK"
      Case 37 : Return "MEDIA SERVER"
      Case 38 : Return "MINI DISK"
      Case 39 : Return "MOOD"
      Case 40 : Return "MORNING MUSIC"
      Case 41 : Return "MP3"
      Case 42 : Return "OLDIES"
      Case 43 : Return "POP"
      Case 44 : Return "REAR DOOR"
      Case 45 : Return "RELIGIOUS"
      Case 46 : Return "REPLAYTV"
      Case 47 : Return "ROCK"
      Case 48 : Return "SATELLITE"
      Case 49 : Return "SATELLITE 1"
      Case 50 : Return "SATELLITE 2"
      Case 51 : Return "SATELLITE 3"
      Case 52 : Return "SPECIAL"
      Case 53 : Return "TAPE"
      Case 54 : Return "TAPE 1"
      Case 55 : Return "TAPE 2"
      Case 56 : Return "TIVO"
      Case 57 : Return "TUNER 1"
      Case 58 : Return "TUNER 2"
      Case 59 : Return "TUNER 3"
      Case 60 : Return "TUNER"
      Case 61 : Return "TV"
      Case 62 : Return "VCR"
      Case 63 : Return "VCR 1"
      Case 64 : Return "VCR 2"
      Case 65 : Return "SOURCE 1"
      Case 66 : Return "SOURCE 2"
      Case 67 : Return "SOURCE 3"
      Case 68 : Return "SOURCE 4"
      Case 69 : Return "SOURCE 5"
      Case 70 : Return "SOURCE 6"
      Case 71 : Return "SOURCE 7"
      Case 72 : Return "SOURCE 8"
      Case 73 : Return "CUSTOM NAME 1 "
      Case 74 : Return "CUSTOM NAME 2"
      Case 75 : Return "CUSTOM NAME 3"
      Case 76 : Return "CUSTOM NAME 4"
      Case 77 : Return "CUSTOM NAME 5"
      Case 78 : Return "CUSTOM NAME 6"
      Case 79 : Return "CUSTOM NAME 7"
      Case 80 : Return "CUSTOM NAME 8"
      Case 81 : Return "CUSTOM NAME 9"
      Case 82 : Return "CUSTOM NAME 10"
      Case Else : Return "UNKNOWN"
    End Select

  End Function

  Private Function GetRussoundMasterString(index As Integer) As String

    Select Case index
      Case 0 : Return "Unassigned"
      Case 1 : Return "#"
      Case 2 : Return "# Of SOURCES"
      Case 3 : Return "0"
      Case 4 : Return "1"
      Case 5 : Return "11"
      Case 6 : Return "12"
      Case 7 : Return "13"
      Case 8 : Return "14"
      Case 9 : Return "15"
      Case 10 : Return "16"
      Case 11 : Return "2"
      Case 12 : Return "3"
      Case 13 : Return "4"
      Case 14 : Return "5"
      Case 15 : Return "6"
      Case 16 : Return "7"
      Case 17 : Return "70 MM"
      Case 18 : Return "8"
      Case 19 : Return "9"
      Case 20 : Return "A / B"
      Case 21 : Return "All Zones"
      Case 22 : Return "AM / FM"
      Case 23 : Return "Amber"
      Case 24 : Return "Ambiance"
      Case 25 : Return "Amp"
      Case 26 : Return "ARE YOU SURE"
      Case 27 : Return "AssignKeypad"
      Case 28 : Return "Audio"
      Case 29 : Return "AudioSensing"
      Case 30 : Return "Auto Play"
      Case 31 : Return "AUTO PLAY"
      Case 32 : Return "AUTO PLAY?"
      Case 33 : Return "AUTO SETUP"
      Case 34 : Return "Aux 1"
      Case 35 : Return "Aux 2"
      Case 36 : Return "Aux"
      Case 37 : Return "BALANCE"
      Case 38 : Return "Bargraph"
      Case 39 : Return "BASIC SETUP"
      Case 40 : Return "BASS"
      Case 41 : Return "Bass -"
      Case 42 : Return "Bass +"
      Case 43 : Return "BackgndColor"
      Case 44 : Return "BG COLOR"
      Case 45 : Return "Blues"
      Case 46 : Return "Both"
      Case 47 : Return "Bright"
      Case 48 : Return "BUILD Date"
      Case 49 : Return "BUILD TIME"
      Case 50 : Return "BUTTON TEST"
      Case 51 : Return "Cable 1"
      Case 52 : Return "Cable 2"
      Case 53 : Return "Cable 3"
      Case 54 : Return "Cable"
      Case 55 : Return "CassetteTape"
      Case 56 : Return "CAV Param"
      Case 57 : Return "CAV PARAM"
      Case 58 : Return "CD"
      Case 59 : Return "CD Changer"
      Case 60 : Return "CD Changer 1"
      Case 61 : Return "CD Changer 2"
      Case 62 : Return "CD Changer 3"
      Case 63 : Return "CD Player"
      Case 64 : Return "CD Player 1"
      Case 65 : Return "CD Player 2"
      Case 66 : Return "CD Player 3"
      Case 67 : Return "Channel"
      Case 68 : Return "Channel Dn"
      Case 69 : Return "Channel Up"
      Case 70 : Return "Classical"
      Case 71 : Return "Close"
      Case 72 : Return "COMMAND NUM"
      Case 73 : Return "CommandPool"
      Case 74 : Return "COMMAND TYPE"
      Case 75 : Return "Computer"
      Case 76 : Return "COPY CONFIG"
      Case 77 : Return "COPY To..."
      Case 78 : Return "CONTROLLR ID"
      Case 79 : Return "Ctrl ID Flags"
      Case 80 : Return "Country"
      Case 81 : Return "CTRLR SETUP"
      Case 82 : Return "Cue"
      Case 83 : Return "CustomName"
      Case 84 : Return "CUSTOM NAMES"
      Case 85 : Return "Dance Music"
      Case 86 : Return "DAT"
      Case 87 : Return "Data"
      Case 88 : Return "Default"
      Case 89 : Return "Dflt Dev Cod"
      Case 90 : Return "Deflt Dev Tp"
      Case 91 : Return "Delay"
      Case 92 : Return "DELAY TIME"
      Case 93 : Return "Delete IR"
      Case 94 : Return "DetectAudio"
      Case 95 : Return "DEVICE CODE"
      Case 96 : Return "DEVICE NAMES"
      Case 97 : Return "DEVICE TYPE"
      Case 98 : Return "Diagnostics"
      Case 99 : Return "DIAGNOSTICS"
      Case 100 : Return "Digtl Cable"
      Case 101 : Return "Dim"
      Case 102 : Return "Disable"
      Case 103 : Return "Disk"
      Case 104 : Return "Disk Down"
      Case 105 : Return "Disk Up"
      Case 106 : Return "Display"
      Case 107 : Return "DISP BKLIGHT"
      Case 108 : Return "DISP BLOCK"
      Case 109 : Return "DISP Char"
      Case 110 : Return "DISP ROW"
      Case 111 : Return "DoNotDisturb"
      Case 112 : Return "Do Not DSTRB"
      Case 113 : Return "Dolby Digitl"
      Case 114 : Return "DTS"
      Case 115 : Return "DSS 1"
      Case 116 : Return "DSS 2"
      Case 117 : Return "DSS 3"
      Case 118 : Return "DSS Receiver"
      Case 119 : Return "DVD"
      Case 120 : Return "DVD Changer"
      Case 121 : Return "DVD Changr 1"
      Case 122 : Return "DVD Changr 2"
      Case 123 : Return "DVD Changr 3"
      Case 124 : Return "DVD Player 1"
      Case 125 : Return "DVD Player 2"
      Case 126 : Return "DVD Player 3"
      Case 127 : Return "DVD Player"
      Case 128 : Return "Enable"
      Case 129 : Return "Enter"
      Case 130 : Return "Error Log"
      Case 131 : Return "EventHandler"
      Case 132 : Return "Exit"
      Case 133 : Return "ExtInterface"
      Case 134 : Return "FACTORY INIT"
      Case 135 : Return "False"
      Case 136 : Return "Fast Fwd"
      Case 137 : Return "Fav/Funct 1"
      Case 138 : Return "Fav/Funct 2"
      Case 139 : Return "Favorite"
      Case 140 : Return "FlashDisplay"
      Case 141 : Return "FormatID"
      Case 142 : Return "Front A/V In"
      Case 143 : Return "FRONT A/V In"
      Case 144 : Return "Front Door"
      Case 145 : Return "Function 1"
      Case 146 : Return "Function 2"
      Case 147 : Return "Green"
      Case 148 : Return "Guide"
      Case 149 : Return "Halt"
      Case 150 : Return "HIGHEST NUM"
      Case 151 : Return "Home Control"
      Case 152 : Return "Info"
      Case 153 : Return "Input"
      Case 154 : Return "IntrnetRadio"
      Case 155 : Return "IR IC Test"
      Case 156 : Return "IR Learning"
      Case 157 : Return "Jazz"
      Case 158 : Return "Key"
      Case 159 : Return "KEY CODE"
      Case 160 : Return "KEY CONFIG"
      Case 161 : Return "KEY Function"
      Case 162 : Return "Key Hold"
      Case 163 : Return "KEY NAME"
      Case 164 : Return "Key Press"
      Case 165 : Return "KEY TYPE"
      Case 166 : Return "KEYPAD Function"
      Case 167 : Return "Keypads"
      Case 168 : Return "Keys"
      Case 169 : Return "Laser Disk"
      Case 170 : Return "Last"
      Case 171 : Return "LEARN/DELETE"
      Case 172 : Return "LEARN IR"
      Case 173 : Return "Learn IR Now"
      Case 174 : Return "Learned Code"
      Case 175 : Return "Lrnd Code ID"
      Case 176 : Return "Learnd Codes"
      Case 177 : Return "Learned IR"
      Case 178 : Return "LEARNED SRC"
      Case 179 : Return "LOUDNESS"
      Case 180 : Return "Lrnd Source"
      Case 181 : Return "Lrnd Sources"
      Case 182 : Return "Macro"
      Case 183 : Return "Macro 1"
      Case 184 : Return "Macro 2"
      Case 185 : Return "Macro 3"
      Case 186 : Return "Macro 4"
      Case 187 : Return "MACRO EDITOR"
      Case 188 : Return "MACRO ID"
      Case 189 : Return "MACRO NAME"
      Case 190 : Return "MACRO NUM"
      Case 191 : Return "Main Trigger"
      Case 192 : Return "Master"
      Case 193 : Return "MASTER ENABL"
      Case 194 : Return "Num Scrl Num"
      Case 195 : Return "MaxLevels"
      Case 196 : Return "MAX VOLUME"
      Case 197 : Return "Media Server"
      Case 198 : Return "Menu"
      Case 199 : Return "Menu Dn"
      Case 200 : Return "Menu Left"
      Case 201 : Return "Menu Right"
      Case 202 : Return "Menu Up"
      Case 203 : Return "Mini Disk"
      Case 204 : Return "Minus"
      Case 205 : Return "Mood Music"
      Case 206 : Return "MorningMusic"
      Case 207 : Return "MP3"
      Case 208 : Return "Mute"
      Case 209 : Return "Next"
      Case 210 : Return "Next Chapter"
      Case 211 : Return "Next Song"
      Case 212 : Return "Next Track"
      Case 213 : Return "No"
      Case 214 : Return "None"
      Case 215 : Return "#CONTROLLERS"
      Case 216 : Return "NumObjectTypes"
      Case 217 : Return "NUM ZONES"
      Case 218 : Return "NUMERIC IR"
      Case 219 : Return "Num Pfx Cmd"
      Case 220 : Return "Num Sfx Cmd"
      Case 221 : Return "NUMERIC TEXT"
      Case 222 : Return "Off"
      Case 223 : Return "Oldies"
      Case 224 : Return "On"
      Case 225 : Return "Open"
      Case 226 : Return "Open / Close"
      Case 227 : Return "Page Down"
      Case 228 : Return "PAGE ENABLE"
      Case 229 : Return "Page Up"
      Case 230 : Return "PAGE VOLUME"
      Case 231 : Return "PARAM VALUE"
      Case 232 : Return "Party"
      Case 233 : Return "PARTY ENABLE"
      Case 234 : Return "PARTY MODE"
      Case 235 : Return "Pause"
      Case 236 : Return "PeekPoke Data"
      Case 237 : Return "PeekPoke Setup"
      Case 238 : Return "Peripheral"
      Case 239 : Return "PIP"
      Case 240 : Return "PIP Move"
      Case 241 : Return "PIP Swap"
      Case 242 : Return "Phonograph"
      Case 243 : Return "Play"
      Case 244 : Return "Plus"
      Case 245 : Return "Plus 10"
      Case 246 : Return "Pop"
      Case 247 : Return "PORT ID"
      Case 248 : Return "Power"
      Case 249 : Return "POWER"
      Case 250 : Return "POWER MGT"
      Case 251 : Return "Power On"
      Case 252 : Return "Pwr Mgt Stat"
      Case 253 : Return "Pwr On Cmd"
      Case 254 : Return "PWR On CMD?"
      Case 255 : Return "Power Off"
      Case 256 : Return "Pwr Off Cmd"
      Case 257 : Return "PWR OFF CMD?"
      Case 258 : Return "PREFIX CMD ?"
      Case 259 : Return "Preset"
      Case 260 : Return "Preset Dn"
      Case 261 : Return "Preset Up"
      Case 262 : Return "Press & Hold"
      Case 263 : Return "Prev Channel"
      Case 264 : Return "Prev Chapter"
      Case 265 : Return "Prev Song"
      Case 266 : Return "Prev Track"
      Case 267 : Return "Previous"
      Case 268 : Return "Product Name"
      Case 269 : Return "Product Specific"
      Case 270 : Return "Pro Logic"
      Case 271 : Return "Program"
      Case 272 : Return "Random"
      Case 273 : Return "Rear Door"
      Case 274 : Return "Recall"
      Case 275 : Return "Record"
      Case 276 : Return "Religious"
      Case 277 : Return "RemoteInface"
      Case 278 : Return "ReplayTV"
      Case 279 : Return "Rewind"
      Case 280 : Return "Rock"
      Case 281 : Return "ROOT MENU"
      Case 282 : Return "RUN MENU"
      Case 283 : Return "Sat / DSS"
      Case 284 : Return "Satellite"
      Case 285 : Return "Satellite 1"
      Case 286 : Return "Satellite 2"
      Case 287 : Return "Satellite 3"
      Case 288 : Return "SAVE CHANGES"
      Case 289 : Return "SAVE To.."
      Case 290 : Return "Search Fwd"
      Case 291 : Return "Search Rev"
      Case 292 : Return "Select"
      Case 293 : Return "Select A KEY"
      Case 294 : Return "SENSE DELAY"
      Case 295 : Return "SenseEnable"
      Case 296 : Return "SenseSource"
      Case 297 : Return "SenseStates"
      Case 298 : Return "Setup"
      Case 299 : Return "SETUP MENU"
      Case 300 : Return "Shared"
      Case 301 : Return "Shuffle"
      Case 302 : Return "Size"
      Case 303 : Return "Sleep"
      Case 304 : Return "Source"
      Case 305 : Return "SOURCE"
      Case 306 : Return "Source 1"
      Case 307 : Return "Source 2"
      Case 308 : Return "Source 3"
      Case 309 : Return "Source 4"
      Case 310 : Return "Source 5"
      Case 311 : Return "Source 6"
      Case 312 : Return "Source 7"
      Case 313 : Return "Source 8"
      Case 314 : Return "Source Name"
      Case 315 : Return "SOURCE NAME"
      Case 316 : Return "SOURCE NAMES"
      Case 317 : Return "SOURCE NUM"
      Case 318 : Return "SOURCE SETUP"
      Case 319 : Return "SRC SEL CMD"
      Case 320 : Return "SRC VOL TRIM"
      Case 321 : Return "Sources"
      Case 322 : Return "Special"
      Case 323 : Return "StartAssgnmt"
      Case 324 : Return "StdInterface"
      Case 325 : Return "Stop"
      Case 326 : Return "Storage"
      Case 327 : Return "SUCCESS?"
      Case 328 : Return "SUFFIX CMD ?"
      Case 329 : Return "Sur Mode"
      Case 330 : Return "Surr Mode 1"
      Case 331 : Return "Surr Mode 2"
      Case 332 : Return "Surr Mode 3"
      Case 333 : Return "Surr Mode 4"
      Case 334 : Return "Surr Mode 5"
      Case 335 : Return "Surr Mode 6"
      Case 336 : Return "Surr Mode 7"
      Case 337 : Return "Surr Mode 8"
      Case 338 : Return "Surr Mode 9"
      Case 339 : Return "Surr Mode 10"
      Case 340 : Return "Sur On/Off"
      Case 341 : Return "Surround Dn"
      Case 342 : Return "Surround Up"
      Case 343 : Return "System Cfg"
      Case 344 : Return "System Info"
      Case 345 : Return "SYSTEM INFO"
      Case 346 : Return "SysOn"
      Case 347 : Return "SYSON ENABLE"
      Case 348 : Return "Tape"
      Case 349 : Return "Tape 1"
      Case 350 : Return "Tape 2"
      Case 351 : Return "TEMPLATE TYPE"
      Case 352 : Return "Terminal Rec"
      Case 353 : Return "TerminalSend"
      Case 354 : Return "TEST IR?"
      Case 355 : Return "Theater"
      Case 356 : Return "This Zone"
      Case 357 : Return "TIVO"
      Case 358 : Return "Trace Log"
      Case 359 : Return "Track Fwd"
      Case 360 : Return "Track Rev"
      Case 361 : Return "TREBLE"
      Case 362 : Return "Treble -"
      Case 363 : Return "Treble +"
      Case 364 : Return "Tree Top"
      Case 365 : Return "TRIM LEVEL"
      Case 366 : Return "True"
      Case 367 : Return "Tuner"
      Case 368 : Return "Tuner 1"
      Case 369 : Return "Tuner 2"
      Case 370 : Return "Tuner 3"
      Case 371 : Return "Tuner / Amp"
      Case 372 : Return "TURN On SRC"
      Case 373 : Return "TURN On VOL"
      Case 374 : Return "TV"
      Case 375 : Return "TV / DSS"
      Case 376 : Return "TV / DVD"
      Case 377 : Return "TV / LD"
      Case 378 : Return "TV / VCR"
      Case 379 : Return "TV / Video"
      Case 380 : Return "UEI"
      Case 381 : Return "Univ.Keypad"
      Case 382 : Return "UPDATE FIRMW"
      Case 383 : Return "USE NUM IR ?"
      Case 384 : Return "USER MENU"
      Case 385 : Return "VCR"
      Case 386 : Return "VCR 1"
      Case 387 : Return "VCR 2"
      Case 388 : Return "VERSION"
      Case 389 : Return "V Acc"
      Case 390 : Return "Volume"
      Case 391 : Return "VOLUME"
      Case 392 : Return "Volume Up"
      Case 393 : Return "Volume Down"
      Case 394 : Return "Working Byte"
      Case 395 : Return "Yes"
      Case 396 : Return "Zone"
      Case 397 : Return "ZONE"
      Case 398 : Return "ZONE NUM"
      Case 399 : Return "Zone On Cmd"
      Case 400 : Return "Zone Off Cmd"
      Case 401 : Return "ZONE SETUP"
      Case 402 : Return "Zone Source"
      Case 403 : Return "Zone Sources"
      Case 404 : Return "Zone Trigger"
      Case 405 : Return "ZON VOL TRIM"
      Case 406 : Return "Zone 1"
      Case 407 : Return "Zone 2"
      Case 408 : Return "Zone 3"
      Case 409 : Return "Zone 4"
      Case 410 : Return "Zone 5"
      Case 411 : Return "Zone 6"
      Case 412 : Return "Zone1 Source"
      Case 413 : Return "Zone2 Source"
      Case 414 : Return "Zone3 Source"
      Case 415 : Return "Zone4 Source"
      Case 416 : Return "Zone5 Source"
      Case 417 : Return "Zone6 Source"
      Case 418 : Return "Zone1 Volume"
      Case 419 : Return "Zone2 Volume"
      Case 420 : Return "Zone3 Volume"
      Case 421 : Return "Zone4 Volume"
      Case 422 : Return "Zone5 Volume"
      Case 423 : Return "Zone6 Volume"
      Case 424 : Return "Zones"
      Case 425 : Return "cust 1"
      Case 426 : Return "cust 2"
      Case 427 : Return "cust 3"
      Case 428 : Return "cust 4"
      Case 429 : Return "cust 5"
      Case 430 : Return "cust 6"
      Case 431 : Return "cust 7"
      Case 432 : Return "cust 8"
      Case 433 : Return "cust 9"
      Case 434 : Return "cust 10"
      Case 435 : Return "External Src"
      Case 436 : Return "Live/Intro"
      Case 437 : Return "Setup Menu"
      Case 438 : Return "Back"
      Case 439 : Return "Fav Channel"
      Case 440 : Return "Display Fmt"
      Case 441 : Return "SAP"
      Case 442 : Return "Slow"
      Case 443 : Return "PIP On"
      Case 444 : Return "PIP Off"
      Case 445 : Return "PIP Freeze"
      Case 446 : Return "PIP Input"
      Case 447 : Return "PIP Chan Up"
      Case 448 : Return "PIP Chan Dn"
      Case 449 : Return "PIP Multi"
      Case 450 : Return "Input 1"
      Case 451 : Return "Input 2"
      Case 452 : Return "Input 3"
      Case 453 : Return "Input 4"
      Case 454 : Return "Input 5"
      Case 455 : Return "Input 6"
      Case 456 : Return "Input 7"
      Case 457 : Return "Input 8"
      Case 458 : Return "Input 9"
      Case 459 : Return "Input 10"
      Case 460 : Return "Zone Info"
      Case 461 : Return "Please Wait"
      Case 462 : Return "Disk Loading"
      Case 463 : Return "DVD Suffix"
      Case 464 : Return "DVD Prefix"
      Case 465 : Return "10"
      Case 466 : Return "Living Room"
      Case 467 : Return "Kitchen"
      Case 468 : Return "Bedroom"
      Case 469 : Return "Bedroom 1"
      Case 470 : Return "Bedroom 2"
      Case 471 : Return "Bedroom 3"
      Case 472 : Return "Bedroom 4"
      Case 473 : Return "Bedroom 5"
      Case 474 : Return "Family Room"
      Case 475 : Return "Den"
      Case 476 : Return "Basement"
      Case 477 : Return "Front Yard"
      Case 478 : Return "Back Yard"
      Case 479 : Return "Deck"
      Case 480 : Return "Bathroom"
      Case 481 : Return "Bathroom 1"
      Case 482 : Return "Bathroom 2"
      Case 483 : Return "Bathroom 3"
      Case 484 : Return "Bathroom 4"
      Case 485 : Return "Garden"
      Case 486 : Return "Pool Area"
      Case 487 : Return "Pool Room"
      Case 488 : Return "Studio"
      Case 489 : Return "Control Room"
      Case 490 : Return "Master Bedroom"
      Case 491 : Return "Dining Room"
      Case 492 : Return "Tennis Court"
      Case 493 : Return "Sauna"
      Case 494 : Return "Office"
      Case 495 : Return "Office 1"
      Case 496 : Return "Office 2"
      Case 497 : Return "Office 3"
      Case 498 : Return "Office 4"
      Case 499 : Return "AMP/RCVR Set"
      Case 500 : Return "SYSTEM"
      Case 501 : Return "USE STATUS?"
      Case 502 : Return "BANK NAME"
      Case 503 : Return "BANK #"
      Case 504 : Return "MEMORY NAME"
      Case 505 : Return "MEMORY #"
      Case 506 : Return "TUNER #"
      Case 507 : Return "REGION"
      Case 508 : Return "US"
      Case 509 : Return "Euro"
      Case Else : Return "Unknown"
    End Select

  End Function

#End Region

  Public Enum ZoneControlType As Byte
    Bass = &H00
    Treble = &H01
    Loudness = &H02
    Balance = &H03
    TurnOnVolume = &H4
    BackgroundColor = &H5
    DoNotDisturb = &H6
    PartyMode = &H7
  End Enum

  Public Enum PacketMessageType As Byte
    SetData = &H00
    RequestData = &H01
    Handshake = &H02
    [Event] = &H05
    Display = &H06
  End Enum

  Public Enum PacketEventType As Byte
    ZoneTone = &H00
    ZoneVolume = &H01
    ZoneSourceInput = &H02
    ZoneAllInfo = &H07
  End Enum

  Public Enum XMCategories
    <Description("None")>
    None = &H00
    <Description("Decades")>
    Decades = &H01
    <Description("Country")>
    Country = &H02
    <Description("Hits")>
    Hits = &H03
    <Description("Christian")>
    Christian = &H04
    <Description("Rock")>
    Rock = &H05
    <Description("Urban")>
    Urban = &H06
    <Description("Jazz & Blues")>
    JazzAndBlues = &H07
    <Description("Lifestyle")>
    Lifestyle = &H08
    <Description("Dance")>
    Dance = &H09
    <Description("Latin")>
    Latin = &H0A
    <Description("World")>
    World = &H0B
    <Description("Classical")>
    Classical = &H0C
    <Description("Kids")>
    Kids = &H0D
    <Description("News")>
    News = &H0E
    <Description("Sports")>
    Sports = &H0F
    <Description("Comedy")>
    Comedy = &H10
    <Description("Talk & Entertainment")>
    TalkAndEntertainment = &H11
    <Description("Special Events")>
    Special_Events = &H12
    <Description("Traffic")>
    Traffic = 13
  End Enum

  Public Enum KeypadEvents
    <Description("Setup Button")>
    Setup_Button = &H64
    <Description("Previous")>
    Previous = &H67
    <Description("Next")>
    [Next] = &H68
    <Description("Plus")>
    Plus = &H69
    <Description("Minus")>
    Minus = &H6A
    <Description("Source Toggle")>
    Source_Toggle = &H6B
    <Description("Power")>
    Power = &H6C
    <Description("Stop")>
    [Stop] = &H6D
    <Description("Pause")>
    Pause = &H6E
    <Description("Favorite 1")>
    Favorite_1 = &H6F
    <Description("Favorite 2")>
    Favorite_2 = &H70
    <Description("Play")>
    Play = &H73
    <Description("Volume Up")>
    Volume_Up = &H7F
    <Description("Volume Down")>
    Volume_Down = &H80
  End Enum

  Public Enum TunerCommand
    <Description("Digit 0")>
    DIGIT_0 = &H0A
    <Description("Digit 1")>
    DIGIT_1 = &H01
    <Description("Digit 2")>
    DIGIT_2 = &H02
    <Description("Digit 3")>
    DIGIT_3 = &H03
    <Description("Digit 4")>
    DIGIT_4 = &H04
    <Description("Digit 5")>
    DIGIT_5 = &H05
    <Description("Digit 6")>
    DIGIT_6 = &H06
    <Description("Digit 7")>
    DIGIT_7 = &H07
    <Description("Digit 8")>
    DIGIT_8 = &H08
    <Description("Digit 9")>
    DIGIT_9 = &H09
    <Description("Preset UP")>
    PRESET_UP = &H0E
    <Description("Preset Down")>
    PRESET_DOWN = &H0F
    <Description("Direct Bank Selection Mode")>
    DIRECT_BANK_SELECTION_MODE = &H15
    <Description("Direct Preset Selection Mode")>
    DIRECT_PRESET_SELECTION_MODE = &H16
    <Description("Direct Tuning Mode")>
    DIRECT_TUNING_MODE = &H17
    <Description("Bank Up")>
    BANK_UP = &H29
    <Description("Bank Down")>
    BANK_DOWN = &H2A
    <Description("Tune Up")>
    TUNE_UP = &H2F
    <Description("Tune Down")>
    TUNE_DOWN = &H30
    <Description("Power On")>
    POWER_ON = &H3A
    <Description("Power Off")>
    POWER_OFF = &H3B
    <Description("Seek")>
    SEEK = &H1B
    <Description("Scan")>
    SCAN = &H46
    <Description("AM/FM Toggle")>
    AM_FM_TOGGLE = &H1A
    <Description("FM Select")>
    FM_SELECT = &H49
    <Description("AM Select")>
    AM_SELECT = &H4A
    <Description("Stereo/Mono Toggle")>
    STEREO_MONO_TOGGLE = &H13
    <Description("Direct Stereo Select")>
    DIRECT_STEREO_SELECT = &H33
    <Description("Direct Mono Select")>
    DIRECT_MONO_SELECT = &H34
    <Description("Local Distant Toggle")>
    LOCAL_DISTANT_TOGGLE = &H14
    <Description("Direct Local Select")>
    DIRECT_LOCAL_SELECT = &H42
    <Description("Direct Distance Select")>
    DIRECT_DISTANT_SELECT = &H43
    <Description("Category Select")>
    CATEGORY_SELECT = &H4D
    <Description("Category Up")>
    CATEGORY_UP = &H54
    <Description("Category Down")>
    CATEGORY_DOWN = &H55
    <Description("Category Channel Up")>
    CAT_CHANNEL_UP = &H58
    <Description("Category Channel Down")>
    CAT_CHANNEL_DOWN = &H59
  End Enum

  Public Enum Keycodes
    <Description("Button 1")>
    Button_1 = &H01
    <Description("Button 2")>
    Button_2 = &H02
    <Description("Button 3")>
    Button_3 = &H03
    <Description("Button 4")>
    Button_4 = &H04
    <Description("Button 5")>
    Button_5 = &H05
    <Description("Button 6")>
    Button_6 = &H06
    <Description("Button 7")>
    Button_7 = &H07
    <Description("Button 8")>
    Button_8 = &H08
    <Description("Button 9")>
    Button_9 = &H09
    <Description("Button 0")>
    Button_0 = &H0A
    <Description("Volume Up")>
    Volume_Up = &H0B
    <Description("Volume Down")>
    Volume_Down = &H0C
    <Description("Mute")>
    Mute = &H0D
    <Description("Channel Up")>
    Channel_Up = &H0E
    <Description("Channel Down")>
    Channel_Down = &H0F
    <Description("Power")>
    Power = &H10
    <Description("Enter")>
    Enter = &H11
    <Description("Previous Channel")>
    Previous_Channel = &H12
    <Description("TV/Video")>
    TV_Video = &H13
    <Description("TV/VCR")>
    TV_VCR = &H14
    <Description("A/B")>
    A_B = &H15
    <Description("TV/DVD")>
    TV_DVD = &H16
    <Description("TV/LD")>
    TV_LD = &H17
    <Description("Input")>
    Input = &H18
    <Description("TV/DSS")>
    TV_DSS = &H19
    <Description("Play")>
    Play = &H1A
    <Description("Stop")>
    [Stop] = &H1B
    <Description("Search Forward")>
    Search_Forward = &H1C
    <Description("Search Rewind")>
    Search_Rewind = &H1D
    <Description("Pause")>
    Pause = &H1E
    <Description("Record")>
    Record = &H1F
    <Description("Menu")>
    Menu = &H20
    <Description("Menu Up")>
    Menu_Up = &H21
    <Description("Menu Down")>
    Menu_Down = &H22
    <Description("Menu Left")>
    Menu_Left = &H23
    <Description("Menu Right")>
    Menu_Right = &H24
    <Description("Select")>
    [Select] = &H25
    <Description("Exit")>
    [Exit] = &H26
    <Description("Display")>
    Display = &H27
    <Description("Guide")>
    Guide = &H28
    <Description("Page Up")>
    Page_Up = &H29
    <Description("Page Down")>
    Page_Down = &H2A
    <Description("Disk")>
    Disk = &H2B
    <Description("Plus")>
    Plus = &H2C
    <Description("Open/Close")>
    Open_Close = &H2D
    <Description("Random")>
    Random = &H2E
    <Description("Track Forward")>
    Track_Forward = &H2F
    <Description("Track Reverse")>
    Track_Reverse = &H30
    <Description("Surround On/Off")>
    Surround_On_Off = &H31
    <Description("Surround Mode")>
    Surround_Mode = &H32
    <Description("Surround Up")>
    Surround_Up = &H33
    <Description("Surround Down")>
    Surround_Down = &H34
    <Description("PIP")>
    PIP = &H35
    <Description("PIP Move")>
    PIP_Move = &H36
    <Description("PIP Swap")>
    PIP_Swap = &H37
    <Description("Program")>
    Program = &H38
    <Description("Sleep")>
    Sleep = &H39
    <Description("On")>
    [On] = &H3A
    <Description("Off")>
    Off = &H3B
    <Description("Button 11")>
    Button_11 = &H3C
    <Description("Button 12")>
    Button_12 = &H3D
    <Description("Button 13")>
    Button_13 = &H3E
    <Description("Button 14")>
    Button_14 = &H3F
    <Description("Button 15")>
    Button_15 = &H40
    <Description("Button 16")>
    Button_16 = &H41
    <Description("Bright")>
    Bright = &H42
    <Description("Dim")>
    [Dim] = &H43
    <Description("Close")>
    Close = &H44
    <Description("Open")>
    Open = &H45
    <Description("Stop 2")>
    Stop_2 = &H46
    <Description("AM/FM ")>
    AM_FM = &H47
    <Description("Cue")>
    Cue = &H48
    <Description("Disk Up")>
    Disk_Up = &H49
    <Description("Disk Down")>
    Disk_Down = &H4A
    <Description("Info")>
    Info = &H4B
  End Enum

End Class
