Imports System.Threading
Imports System.Text.RegularExpressions
Imports System.Text
Imports HomeSeerAPI
Imports Scheduler
Imports System.ComponentModel
Imports System.Data.Common
Imports System.Data.SQLite

Module hspi_plugin

  '
  ' Declare public objects, not required by HomeSeer
  '
  Dim actions As New hsCollection
  Dim triggers As New hsCollection
  Dim conditions As New Hashtable
  Const Pagename = "Events"

  Public gAudioDevices As New List(Of hspi_audio_device)
  Public gAudioDeviceLock As New Object

  Public Const IFACE_NAME As String = "UltraRussoundAudio3"

  Public Const LINK_TARGET As String = ""
  Public Const LINK_TEXT As String = "UltraRussoundAudio3"
  Public Const LINK_PAGE_TITLE As String = "UltraRussoundAudio3 HSPI"
  Public Const LINK_HELP As String = "/hspi_ultrarussoundaudio3/UltraRussoundAudio3_HSPI_Users_Guide.pdf"

  Public gBaseCode As String = ""
  Public gIOEnabled As Boolean = True
  Public gImageDir As String = "/images/hspi_ultrarussoundaudio3/"
  Public gHSInitialized As Boolean = False
  Public gINIFile As String = "hspi_" & IFACE_NAME.ToLower & ".ini"

  Public HSAppPath As String = ""

#Region "HSPI - Public Routines"

  ''' <summary>
  ''' Connects to the Audio Device units installed on the network
  ''' </summary>
  ''' <remarks></remarks>
  Public Sub AudioDeviceConnection()

    Dim AudioDeviceList As New SortedList

    Dim bAbortThread As Boolean = False

    Try
      WriteMessage("The Audio Device connection routine has started ...", MessageType.Debug)

      While bAbortThread = False

        Try

          AudioDeviceList = GetAudioDeviceList()
          For Each device_id As Object In AudioDeviceList.Keys
            AddAudioDevice(AudioDeviceList(device_id))
          Next

        Catch pEx As Exception
          '
          ' Return message
          '
          ProcessError(pEx, "AudioDeviceConnection()")
        End Try

        '
        ' Give up some time
        '
        Thread.Sleep(1000 * 60)

      End While ' Stay in thread until we get an abort/exit request

    Catch pEx As ThreadAbortException
      ' 
      ' There was a normal request to terminate the thread.  
      '
      bAbortThread = True      ' Not actually needed
      WriteMessage(String.Format("DiscoveryConnection thread received abort request, terminating normally."), MessageType.Informational)

    Catch pEx As Exception
      '
      ' Return message
      '
      ProcessError(pEx, "AudioDeviceConnection()")

    Finally
      '
      ' Notify that we are exiting the thread
      '
      WriteMessage(String.Format("AudioDeviceConnection terminated."), MessageType.Debug)

    End Try

  End Sub

  ''' <summary>
  ''' Add the Audio Device to the AudioDevices hashtable
  ''' </summary>
  ''' <param name="AudioDevice"></param>
  ''' <remarks></remarks>
  Private Sub AddAudioDevice(ByVal AudioDevice As Hashtable)

    Try

      If AudioDevice.ContainsKey("device_id") = False Then Exit Sub

      If gAudioDevices.Any(Function(s) s.DeviceId = AudioDevice("device_id")) = False Then
        '
        ' This is a new Audio Device
        '
        SyncLock gAudioDeviceLock

          Dim NewAudioDevice As New hspi_audio_device(AudioDevice("device_id"),
                                                      AudioDevice("device_name"),
                                                      AudioDevice("device_serial"),
                                                      AudioDevice("device_conn"),
                                                      AudioDevice("device_addr"),
                                                      AudioDevice("device_make"),
                                                      AudioDevice("device_model"),
                                                      AudioDevice("device_zones"),
                                                      AudioDevice("device_tuner_src"))

          '
          ' Attempt a connection to the audio device
          '
          Dim strResult As String = NewAudioDevice.ConnectToDevice()

          '
          ' Get the Audio Device status
          '
          Dim bConnected As Boolean = NewAudioDevice.CheckDeviceConnection()
          If bConnected = True Then
            WriteMessage(String.Format("Audio Device connection established to {0}.", AudioDevice("device_addr")), MessageType.Informational)

            If NewAudioDevice.DeviceTunerSource >= 0 Then

            End If

            WriteMessage(String.Format("Adding Audio Device Object with Id {0}", AudioDevice("device_id")), MessageType.Debug)
            gAudioDevices.Add(NewAudioDevice)
          End If

        End SyncLock

      End If

    Catch pEx As Exception
      '
      ' Process program exception
      '
      Call ProcessError(pEx, "AddAudioDevice()")
    End Try

  End Sub

  ''' <summary>
  ''' Initialize the Hash Tables
  ''' </summary>
  ''' <remarks></remarks>
  Public Sub InitializeHashTables()

  End Sub

  ''' <summary>
  ''' Get the Audio Device devices
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetAudioDeviceList() As SortedList

    Dim AudioDevices As New SortedList

    Try
      '
      ' Define the SQL Query
      '
      Dim strSQL As String = String.Format("SELECT device_id, device_name, device_serial, device_make, device_model, device_conn, device_addr, device_zones, device_tuner_src FROM tblAudioDevices WHERE device_conn <> '{0}'", "Disabled")

      '
      ' Execute the data reader
      '
      Using MyDbCommand As DbCommand = DBConnectionMain.CreateCommand()

        MyDbCommand.Connection = DBConnectionMain
        MyDbCommand.CommandType = CommandType.Text
        MyDbCommand.CommandText = strSQL

        SyncLock SyncLockMain
          Dim dtrResults As IDataReader = MyDbCommand.ExecuteReader()

          '
          ' Process the resutls
          '
          While dtrResults.Read()
            Dim AudioDevice As New Hashtable

            AudioDevice.Add("device_id", dtrResults("device_id"))
            AudioDevice.Add("device_name", dtrResults("device_name"))
            AudioDevice.Add("device_serial", dtrResults("device_serial"))
            AudioDevice.Add("device_make", dtrResults("device_make"))
            AudioDevice.Add("device_model", dtrResults("device_model"))
            AudioDevice.Add("device_conn", dtrResults("device_conn"))
            AudioDevice.Add("device_addr", dtrResults("device_addr"))
            AudioDevice.Add("device_zones", dtrResults("device_zones"))
            AudioDevice.Add("device_tuner_src", dtrResults("device_tuner_src"))

            If AudioDevices.ContainsKey(dtrResults("device_id")) = False Then
              AudioDevices.Add(dtrResults("device_id"), AudioDevice)
            End If
          End While

          dtrResults.Close()
        End SyncLock

        MyDbCommand.Dispose()

      End Using

    Catch pEx As Exception
      '
      ' Process the error
      '
      Call ProcessError(pEx, "GetAudioDeviceList()")
    End Try

    Return AudioDevices

  End Function

  ''' <summary>
  ''' Get the Audio Device
  ''' </summary>
  ''' <param name="device_id"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetAudioDevice(ByVal device_id As String) As Hashtable

    Dim AudioDevice As New Hashtable

    Try
      '
      ' Define the SQL Query
      '
      Dim strSQL As String = String.Format("SELECT device_id, device_name, device_serial, device_make, device_model, device_conn, device_addr, device_zones, device_tuner_src FROM tblAudioDevices WHERE device_id='{0}'", device_id)

      '
      ' Execute the data reader
      '
      Using MyDbCommand As DbCommand = DBConnectionMain.CreateCommand()

        MyDbCommand.Connection = DBConnectionMain
        MyDbCommand.CommandType = CommandType.Text
        MyDbCommand.CommandText = strSQL

        SyncLock SyncLockMain

          Dim dtrResults As IDataReader = MyDbCommand.ExecuteReader()

          '
          ' Process the resutls
          '
          If dtrResults.Read() Then
            AudioDevice.Add("device_id", dtrResults("device_id"))
            AudioDevice.Add("device_name", dtrResults("device_name"))
            AudioDevice.Add("device_serial", dtrResults("device_serial"))
            AudioDevice.Add("device_make", dtrResults("device_make"))
            AudioDevice.Add("device_model", dtrResults("device_model"))
            AudioDevice.Add("device_conn", dtrResults("device_conn"))
            AudioDevice.Add("device_addr", dtrResults("device_addr"))
            AudioDevice.Add("device_zones", dtrResults("device_zones"))
            AudioDevice.Add("device_tuner_src", dtrResults("device_tuner_src"))
          End If

          dtrResults.Close()
        End SyncLock

        MyDbCommand.Dispose()
      End Using

    Catch pEx As Exception
      '
      ' Process the error
      '
      Call ProcessError(pEx, "GetAudioDevice()")
    End Try

    Return AudioDevice

  End Function

  ''' <summary>
  ''' Gets the Audio Controller Devices from the underlying database
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetAudioDevices() As DataTable

    Dim ResultsDT As New DataTable
    Dim strMessage As String = ""

    strMessage = "Entered GetAudioDevices() function."
    Call WriteMessage(strMessage, MessageType.Debug)

    Try
      '
      ' Make sure the datbase is open before attempting to use it
      '
      Select Case DBConnectionMain.State
        Case ConnectionState.Broken, ConnectionState.Closed
          strMessage = "Unable to complete database query because the database " _
                     & "connection has not been initialized."
          Throw New System.Exception(strMessage)
      End Select

      Dim strSQL As String = String.Format("SELECT * FROM tblAudioDevices")

      '
      ' Initialize the command object
      '
      Dim MyDbCommand As DbCommand = DBConnectionMain.CreateCommand()

      MyDbCommand.Connection = DBConnectionMain
      MyDbCommand.CommandType = CommandType.Text
      MyDbCommand.CommandText = strSQL

      '
      ' Initialize the dataset, then populate it
      '
      Dim MyDS As DataSet = New DataSet

      Dim MyDA As System.Data.IDbDataAdapter = New SQLiteDataAdapter(MyDbCommand)
      MyDA.SelectCommand = MyDbCommand

      SyncLock SyncLockMain
        MyDA.Fill(MyDS)
      End SyncLock

      '
      ' Get our DataTable
      '
      Dim MyDT As DataTable = MyDS.Tables(0)

      '
      ' Get record count
      '
      Dim iRecordCount As Integer = MyDT.Rows.Count

      If iRecordCount > 0 Then
        '
        ' Build field names
        '
        Dim iFieldCount As Integer = MyDS.Tables(0).Columns.Count() - 1
        For iFieldNum As Integer = 0 To iFieldCount
          '
          ' Create the columns
          '
          Dim ColumnName As String = MyDT.Columns.Item(iFieldNum).ColumnName
          Dim MyDataColumn As New DataColumn(ColumnName, GetType(String))

          '
          ' Add the columns to the DataTable's Columns collection
          '
          ResultsDT.Columns.Add(MyDataColumn)
        Next

        '
        ' Let's output our records	
        '
        Dim i As Integer = 0
        For i = 0 To iRecordCount - 1
          '
          ' Create the rows
          '
          Dim dr As DataRow
          dr = ResultsDT.NewRow()
          For iFieldNum As Integer = 0 To iFieldCount
            dr(iFieldNum) = MyDT.Rows(i)(iFieldNum)
          Next
          ResultsDT.Rows.Add(dr)
        Next

      End If

    Catch pEx As Exception
      '
      ' Process Exception
      '
      Call ProcessError(pEx, "GetAudioDevices()")

    End Try

    Return ResultsDT

  End Function

  ''' <summary>
  ''' Inserts a new Audio Device into the database
  ''' </summary>
  ''' <param name="device_name"></param>
  ''' <param name="device_serial"></param>
  ''' <param name="device_make"></param>
  ''' <param name="device_model"></param>
  ''' <param name="device_conn"></param>
  ''' <param name="device_addr"></param>
  ''' <param name="device_zones"></param>
  ''' <param name="device_tuner_src"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function InsertAudioDevice(ByVal device_name As String,
                                    ByVal device_serial As String,
                                    ByVal device_make As String,
                                    ByVal device_model As String,
                                    ByVal device_conn As String,
                                    ByVal device_addr As String,
                                    ByVal device_zones As Integer,
                                    ByVal device_tuner_src As Integer,
                                    ByVal device_media_mgr As Integer) As Integer

    Dim strMessage As String = ""
    Dim iRecordsAffected As Integer = 0

    Try

      Select Case DBConnectionMain.State
        Case ConnectionState.Broken, ConnectionState.Closed
          strMessage = "Unable to complete database transaction because the database " _
                     & "connection has not been initialized."
          Throw New System.Exception(strMessage)
      End Select

      If device_serial.Length = 0 Then
        Throw New Exception("One or more required fields are empty.  Unable to insert new Audio Device into the database.")
      ElseIf device_conn = "Ethernet" And (device_addr.Length = 0) Then
        Throw New Exception("The IP address and port fields are required.  Unable to insert new Audio Device into the database.")
      End If

      '
      ' Try inserting the Audio Device into one of the 10 available slots
      '
      For device_id As Integer = 1 To 10

        Dim strSQL As String = String.Format("INSERT INTO tblAudioDevices (" _
                                     & " device_id, device_name, device_serial, device_make, device_model, device_conn, device_addr, device_zones, device_tuner_src, device_media_mgr" _
                                     & ") VALUES (" _
                                     & " {0}, '{1}', '{2}', '{3}', '{4}', '{5}', '{6}', {7}, {8}, {9}" _
                                     & ")", device_id, device_name, device_serial, device_make, device_model, device_conn, device_addr, device_zones, device_tuner_src, device_media_mgr)

        Using dbcmd As DbCommand = DBConnectionMain.CreateCommand()

          dbcmd.Connection = DBConnectionMain
          dbcmd.CommandType = CommandType.Text
          dbcmd.CommandText = strSQL

          Try

            SyncLock SyncLockMain
              iRecordsAffected = dbcmd.ExecuteNonQuery()
            End SyncLock

          Catch pEx As Exception
            '
            ' Ignore this error
            '
          Finally
            dbcmd.Dispose()
          End Try

          If iRecordsAffected > 0 Then
            Return device_id
          End If

        End Using

      Next

      Throw New Exception("Unable to insert Audio Device into the database.  Please ensure you are not attempting to connect more than 10 Audio Devices to the plug-in.")

    Catch pEx As Exception
      Call ProcessError(pEx, "InsertAudioDevice()")
      Return 0
    End Try

  End Function

  ''' <summary>
  ''' Updates existing Audio Device stored in the database
  ''' </summary>
  ''' <param name="device_id"></param>
  ''' <param name="device_name"></param>
  ''' <param name="device_serial"></param>
  ''' <param name="device_make"></param>
  ''' <param name="device_model"></param>
  ''' <param name="device_conn"></param>
  ''' <param name="device_addr"></param>
  ''' <param name="device_zones"></param>
  ''' <param name="device_tuner_src"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function UpdateAudioDevice(ByVal device_id As Integer,
                                   ByVal device_name As String,
                                   ByVal device_serial As String,
                                   ByVal device_make As String,
                                   ByVal device_model As String,
                                   ByVal device_conn As String,
                                   ByVal device_addr As String,
                                   ByVal device_zones As Integer,
                                   ByVal device_tuner_src As Integer,
                                   ByVal device_media_mgr As Integer) As Boolean

    Dim strMessage As String = ""

    Try

      Select Case DBConnectionMain.State
        Case ConnectionState.Broken, ConnectionState.Closed
          strMessage = "Unable to complete database transaction because the database " _
                     & "connection has not been initialized."
          Throw New System.Exception(strMessage)
      End Select

      If device_serial.Length = 0 Then
        Throw New Exception("One or more required fields are empty.  Unable to modify the Audio Device.")
      ElseIf device_conn = "Ethernet" And (device_addr.Length = 0) Then
        Throw New Exception("One or more required fields are empty.  Unable to modify the Audio Device.")
      End If

      Dim strSql As String = String.Format("UPDATE tblAudioDevices SET " _
                                          & " device_name='{0}', " _
                                          & " device_serial='{1}', " _
                                          & " device_make='{2}'," _
                                          & " device_model='{3}'," _
                                          & " device_conn='{4}'," _
                                          & " device_addr='{5}'," _
                                          & " device_zones={6}, " _
                                          & " device_tuner_src={7}, " _
                                          & " device_media_mgr={8} " _
                                          & "WHERE device_id={9}",
                                             device_name,
                                             device_serial,
                                             device_make,
                                             device_model,
                                             device_conn,
                                             device_addr,
                                             device_zones,
                                             device_tuner_src,
                                             device_media_mgr,
                                             device_id.ToString)

      '
      ' Build the insert/update/delete query
      '
      Dim MyDbCommand As DbCommand = DBConnectionMain.CreateCommand()

      MyDbCommand.Connection = DBConnectionMain
      MyDbCommand.CommandType = CommandType.Text
      MyDbCommand.CommandText = strSql

      Dim iRecordsAffected As Integer = 0
      SyncLock SyncLockMain
        iRecordsAffected = MyDbCommand.ExecuteNonQuery()
      End SyncLock

      strMessage = "UpdateAudioDevice() updated " & iRecordsAffected & " row(s)."
      Call WriteMessage(strMessage, MessageType.Debug)

      MyDbCommand.Dispose()

      If iRecordsAffected > 0 Then
        Return True
      Else
        Return False
      End If

    Catch pEx As Exception
      Call ProcessError(pEx, "UpdateAudioDevice()")
      Return False
    End Try

  End Function

  ''' <summary>
  ''' Removes existing Audio Device stored in the database
  ''' </summary>
  ''' <param name="device_id"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function DeleteAudioDevice(ByVal device_id As Integer) As Boolean

    Dim strMessage As String = ""

    Try

      Select Case DBConnectionMain.State
        Case ConnectionState.Broken, ConnectionState.Closed
          strMessage = "Unable to complete database transaction because the database " _
                     & "connection has not been initialized."
          Throw New System.Exception(strMessage)
      End Select

      '
      ' Build the insert/update/delete query
      '
      Dim MyDbCommand As DbCommand = DBConnectionMain.CreateCommand()

      MyDbCommand.Connection = DBConnectionMain
      MyDbCommand.CommandType = CommandType.Text
      MyDbCommand.CommandText = String.Format("DELETE FROM tblAudioDevices WHERE device_id={0}", device_id.ToString)

      Dim iRecordsAffected As Integer = 0
      SyncLock SyncLockMain
        iRecordsAffected = MyDbCommand.ExecuteNonQuery()
      End SyncLock

      strMessage = "DeleteAudioDevice() removed " & iRecordsAffected & " row(s)."
      Call WriteMessage(strMessage, MessageType.Debug)

      MyDbCommand.Dispose()

      If iRecordsAffected > 0 Then
        Return True
      Else
        Return False
      End If

      Return True

    Catch pEx As Exception
      Call ProcessError(pEx, "DeleteAudioDevice()")
      Return False
    End Try

  End Function

#End Region

#Region "HSPI - Misc"

  ''' <summary>
  ''' Gets plug-in setting from INI file
  ''' </summary>
  ''' <param name="strSection"></param>
  ''' <param name="strKey"></param>
  ''' <param name="strValueDefault"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetSetting(ByVal strSection As String,
                             ByVal strKey As String,
                             ByVal strValueDefault As String) As String

    Dim strMessage As String = ""

    Try
      strMessage = "Entered GetSetting() function."
      Call WriteMessage(strMessage, MessageType.Debug)

      '
      ' Get the ini settings
      '
      Dim strValue As String = hs.GetINISetting(strSection, strKey, strValueDefault, gINIFile)

      strMessage = String.Format("Section: {0}, Key: {1}, Value: {2}", strSection, strKey, strValue)
      Call WriteMessage(strMessage, MessageType.Debug)

      '
      ' Check to see if we need to decrypt the data
      '
      If strKey = "UserPass" Then
        strValue = hs.DecryptString(strValue, "&Cul8r#1")
      End If

      Return strValue

    Catch pEx As Exception
      '
      ' Process the error
      '
      Call ProcessError(pEx, "GetSetting()")
      Return ""
    End Try

  End Function

  ''' <summary>
  ''' Saves plug-in setting to INI file
  ''' </summary>
  ''' <param name="strSection"></param>
  ''' <param name="strKey"></param>
  ''' <param name="strValue"></param>
  ''' <remarks></remarks>
  Public Sub SaveSetting(ByVal strSection As String,
                         ByVal strKey As String,
                         ByVal strValue As String)

    Dim strMessage As String = ""

    Try
      strMessage = "Entered SaveSetting() subroutine."
      Call WriteMessage(strMessage, MessageType.Debug)

      '
      ' Check to see if we need to encrypt the data
      '
      If strKey = "UserPass" Then
        If strValue.Length = 0 Then Exit Sub
        strValue = hs.EncryptString(strValue, "&Cul8r#1")
      End If

      strMessage = String.Format("Section: {0}, Key: {1}, Value: {2}", strSection, strKey, strValue)
      Call WriteMessage(strMessage, MessageType.Debug)

      '
      ' Save selected settings to global variables
      '
      'If strSection = "Options" And strKey = "MaxDeliveryAttempts" Then
      '  If IsNumeric(strValue) Then
      '    gMaxAttempts = CInt(Val(strValue))
      '  End If
      'End If

      '
      ' Save the settings
      '
      hs.SaveINISetting(strSection, strKey, strValue, gINIFile)
    Catch pEx As Exception
      '
      ' Process the error
      '
      Call ProcessError(pEx, "SaveSetting()")
    End Try

  End Sub

#End Region

#Region "UltraRussoundAudio3 Actions/Triggers/Conditions"

#Region "Trigger Proerties"

  ''' <summary>
  ''' Defines the valid triggers for this plug-in
  ''' </summary>
  ''' <remarks></remarks>
  Sub SetTriggers()
    Dim o As Object = Nothing
    If triggers.Count = 0 Then
      'triggers.Add(o, "Email Delivery Status")           ' 1
    End If
  End Sub

  ''' <summary>
  ''' Lets HomeSeer know our plug-in has triggers
  ''' </summary>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property HasTriggers() As Boolean
    Get
      SetTriggers()
      Return IIf(triggers.Count > 0, True, False)
    End Get
  End Property

  ''' <summary>
  ''' Returns the trigger count
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function TriggerCount() As Integer
    SetTriggers()
    Return triggers.Count
  End Function

  ''' <summary>
  ''' Returns the subtrigger count
  ''' </summary>
  ''' <param name="TriggerNumber"></param>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property SubTriggerCount(ByVal TriggerNumber As Integer) As Integer
    Get
      Dim trigger As trigger
      If ValidTrig(TriggerNumber) Then
        trigger = triggers(TriggerNumber - 1)
        If Not (trigger Is Nothing) Then
          Return 0
        Else
          Return 0
        End If
      Else
        Return 0
      End If
    End Get
  End Property

  ''' <summary>
  ''' Returns the trigger name
  ''' </summary>
  ''' <param name="TriggerNumber"></param>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property TriggerName(ByVal TriggerNumber As Integer) As String
    Get
      If Not ValidTrig(TriggerNumber) Then
        Return ""
      Else
        Return String.Format("{0}: {1}", IFACE_NAME, triggers.Keys(TriggerNumber - 1))
      End If
    End Get
  End Property

  ''' <summary>
  ''' Returns the subtrigger name
  ''' </summary>
  ''' <param name="TriggerNumber"></param>
  ''' <param name="SubTriggerNumber"></param>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property SubTriggerName(ByVal TriggerNumber As Integer, ByVal SubTriggerNumber As Integer) As String
    Get
      Dim trigger As trigger
      If ValidSubTrig(TriggerNumber, SubTriggerNumber) Then
        Return ""
      Else
        Return ""
      End If
    End Get
  End Property

  ''' <summary>
  ''' Determines if a trigger is valid
  ''' </summary>
  ''' <param name="TrigIn"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Friend Function ValidTrig(ByVal TrigIn As Integer) As Boolean
    SetTriggers()
    If TrigIn > 0 AndAlso TrigIn <= triggers.Count Then
      Return True
    End If
    Return False
  End Function

  ''' <summary>
  ''' Determines if the trigger is a valid subtrigger
  ''' </summary>
  ''' <param name="TrigIn"></param>
  ''' <param name="SubTrigIn"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function ValidSubTrig(ByVal TrigIn As Integer, ByVal SubTrigIn As Integer) As Boolean
    Return False
  End Function

  ''' <summary>
  ''' Tell HomeSeer which triggers have conditions
  ''' </summary>
  ''' <param name="TriggerNumber"></param>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property HasConditions(ByVal TriggerNumber As Integer) As Boolean
    Get
      Select Case TriggerNumber
        Case 0
          Return True   ' Render trigger as IF / AND IF
        Case Else
          Return False  ' Render trigger as IF / OR IF
      End Select
    End Get
  End Property

  ''' <summary>
  ''' HomeSeer will set this to TRUE if the trigger is being used as a CONDITION.  
  ''' Check this value in BuildUI and other procedures to change how the trigger is rendered if it is being used as a condition or a trigger.
  ''' </summary>
  ''' <param name="TrigInfo"></param>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Property Condition(ByVal TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As Boolean
    Set(ByVal value As Boolean)

      Dim UID As String = TrigInfo.UID.ToString

      Dim trigger As New trigger
      If Not (TrigInfo.DataIn Is Nothing) Then
        DeSerializeObject(TrigInfo.DataIn, trigger)
      End If

      ' TriggerCondition(sKey) = value

    End Set
    Get

      Dim UID As String = TrigInfo.UID.ToString

      Dim trigger As New trigger
      If Not (TrigInfo.DataIn Is Nothing) Then
        DeSerializeObject(TrigInfo.DataIn, trigger)
      End If

      Return False

    End Get
  End Property

  ''' <summary>
  ''' Determines if a trigger is a condition
  ''' </summary>
  ''' <param name="sKey"></param>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Property TriggerCondition(sKey As String) As Boolean
    Get

      If conditions.ContainsKey(sKey) = True Then
        Return conditions(sKey)
      Else
        Return False
      End If

    End Get
    Set(value As Boolean)

      If conditions.ContainsKey(sKey) = False Then
        conditions.Add(sKey, value)
      Else
        conditions(sKey) = value
      End If

    End Set
  End Property

  ''' <summary>
  ''' Called when HomeSeer wants to check if a condition is true
  ''' </summary>
  ''' <param name="TrigInfo"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function TriggerTrue(ByVal TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As Boolean

    Dim UID As String = TrigInfo.UID.ToString

    Dim trigger As New trigger
    If Not (TrigInfo.DataIn Is Nothing) Then
      DeSerializeObject(TrigInfo.DataIn, trigger)
    End If

    Return False
  End Function

#End Region

#Region "Trigger Interface"

  ''' <summary>
  ''' Builds the Trigger UI for display on the HomeSeer events page
  ''' </summary>
  ''' <param name="sUnique"></param>
  ''' <param name="TrigInfo"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function TriggerBuildUI(ByVal sUnique As String, ByVal TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As String

    Dim UID As String = TrigInfo.UID.ToString
    Dim stb As New StringBuilder

    Dim trigger As New trigger
    If Not (TrigInfo.DataIn Is Nothing) Then
      DeSerializeObject(TrigInfo.DataIn, trigger)
    Else 'new event, so clean out the trigger object
      trigger = New trigger
    End If

    Select Case TrigInfo.TANumber
      Case AudioTriggers.SomeTrigger
        Dim triggerName As String = GetEnumName(AudioTriggers.SomeTrigger)

        Dim ActionSelected As String = trigger.Item("DeliveryStatus")

        Dim actionId As String = String.Format("{0}{1}_{2}_{3}", triggerName, "DeliveryStatus", UID, sUnique)

        Dim jqDSN As New clsJQuery.jqDropList(actionId, Pagename, True)
        jqDSN.autoPostBack = True

        jqDSN.AddItem("(Select Delivery Status)", "", (ActionSelected = ""))
        Dim Actions As String() = {"Success", "Deferral", "Failure"}
        For Each strAction As String In Actions
          Dim strOptionValue As String = strAction
          Dim strOptionName As String = strOptionValue
          jqDSN.AddItem(strOptionName, strOptionValue, (ActionSelected = strOptionValue))
        Next

        stb.Append(jqDSN.Build)

    End Select

    Return stb.ToString
  End Function

  ''' <summary>
  ''' Process changes to the trigger from the HomeSeer events page
  ''' </summary>
  ''' <param name="PostData"></param>
  ''' <param name="TrigInfo"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function TriggerProcessPostUI(ByVal PostData As System.Collections.Specialized.NameValueCollection,
                                       ByVal TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As HomeSeerAPI.IPlugInAPI.strMultiReturn

    Dim Ret As New HomeSeerAPI.IPlugInAPI.strMultiReturn

    Dim UID As String = TrigInfo.UID.ToString
    Dim TANumber As Integer = TrigInfo.TANumber

    ' When plug-in calls such as ...BuildUI, ...ProcessPostUI, or ...FormatUI are called and there is
    ' feedback or an error condition that needs to be reported back to the user, this string field 
    ' can contain the message to be displayed to the user in HomeSeer UI.  This field is cleared by
    ' HomeSeer after it is displayed to the user.
    Ret.sResult = ""

    ' We cannot be passed info ByRef from HomeSeer, so turn right around and return this same value so that if we want, 
    '   we can exit here by returning 'Ret', all ready to go.  If in this procedure we need to change DataOut or TrigInfo,
    '   we can still do that.
    Ret.DataOut = TrigInfo.DataIn
    Ret.TrigActInfo = TrigInfo

    If PostData Is Nothing Then Return Ret
    If PostData.Count < 1 Then Return Ret

    ' DeSerializeObject
    Dim trigger As New trigger
    If Not (TrigInfo.DataIn Is Nothing) Then
      DeSerializeObject(TrigInfo.DataIn, trigger)
    End If
    trigger.uid = UID

    Dim parts As Collections.Specialized.NameValueCollection = PostData

    Try

      Select Case TANumber
        Case AudioTriggers.SomeTrigger
          Dim triggerName As String = GetEnumName(AudioTriggers.SomeTrigger)

          For Each sKey As String In parts.Keys
            If sKey Is Nothing Then Continue For
            If String.IsNullOrEmpty(sKey.Trim) Then Continue For

            Select Case True
              Case InStr(sKey, triggerName & "DeliveryStatus_" & UID) > 0
                Dim ActionValue As String = parts(sKey)
                trigger.Item("DeliveryStatus") = ActionValue

            End Select
          Next

      End Select

      ' The serialization data for the plug-in object cannot be 
      ' passed ByRef which means it can be passed only one-way through the interface to HomeSeer.
      ' If the plug-in receives DataIn, de-serializes it into an object, and then makes a change 
      ' to the object, this is where the object can be serialized again and passed back to HomeSeer
      ' for storage in the HomeSeer database.

      ' SerializeObject
      If Not SerializeObject(trigger, Ret.DataOut) Then
        Ret.sResult = IFACE_NAME & " Error, Serialization failed. Signal Trigger not added."
        Return Ret
      End If

    Catch ex As Exception
      Ret.sResult = "ERROR, Exception in Trigger UI of " & IFACE_NAME & ": " & ex.Message
      Return Ret
    End Try

    ' All OK
    Ret.sResult = ""
    Return Ret

  End Function

  ''' <summary>
  ''' Determines if a trigger is properly configured
  ''' </summary>
  ''' <param name="TrigInfo"></param>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property TriggerConfigured(ByVal TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As Boolean
    Get
      Dim Configured As Boolean = True
      Dim UID As String = TrigInfo.UID.ToString

      Dim trigger As New trigger
      If Not (TrigInfo.DataIn Is Nothing) Then
        DeSerializeObject(TrigInfo.DataIn, trigger)
      End If

      Select Case TrigInfo.TANumber
        Case AudioTriggers.SomeTrigger
          If trigger.Item("DeliveryStatus") = "" Then Configured = False

      End Select

      Return Configured
    End Get
  End Property

  ''' <summary>
  ''' Formats the trigger for display
  ''' </summary>
  ''' <param name="TrigInfo"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function TriggerFormatUI(ByVal TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As String

    Dim stb As New StringBuilder

    Dim UID As String = TrigInfo.UID.ToString

    Dim trigger As New trigger
    If Not (TrigInfo.DataIn Is Nothing) Then
      DeSerializeObject(TrigInfo.DataIn, trigger)
    End If

    Select Case TrigInfo.TANumber
      Case AudioTriggers.SomeTrigger
        If trigger.uid <= 0 Then
          stb.Append("Trigger has not been properly configured.")
        Else
          Dim strTriggerName As String = GetEnumDescription(AudioTriggers.SomeTrigger)
          Dim strDeliveryStatus As String = trigger.Item("DeliveryStatus")

          stb.AppendFormat("{0} is <font class='event_Txt_Option'>{1}</font>", strTriggerName, strDeliveryStatus)
        End If

    End Select

    Return stb.ToString
  End Function

  ''' <summary>
  ''' Checks to see if trigger should fire
  ''' </summary>
  ''' <param name="Plug_Name"></param>
  ''' <param name="TrigID"></param>
  ''' <param name="SubTrig"></param>
  ''' <param name="strTrigger"></param>
  ''' <remarks></remarks>
  Private Sub CheckTrigger(Plug_Name As String, TrigID As Integer, SubTrig As Integer, strTrigger As String)

    Try
      '
      ' Check HomeSeer Triggers
      '
      If Plug_Name.Contains(":") = False Then Plug_Name &= ":"
      Dim TrigsToCheck() As IAllRemoteAPI.strTrigActInfo = callback.TriggerMatches(Plug_Name, TrigID, SubTrig)

      Try

        For Each TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo In TrigsToCheck
          Dim UID As String = TrigInfo.UID.ToString

          If Not (TrigInfo.DataIn Is Nothing) Then

            Dim trigger As New trigger
            DeSerializeObject(TrigInfo.DataIn, trigger)

            Select Case TrigID

              Case AudioTriggers.SomeTrigger
                Dim strTriggerName As String = GetEnumDescription(AudioTriggers.SomeTrigger)
                Dim strDeliveryStatus As String = trigger.Item("DeliveryStatus")

                Dim strTriggerCheck As String = String.Format("{0},{1}", strTriggerName, strDeliveryStatus)
                If Regex.IsMatch(strTrigger, strTriggerCheck) = True Then
                  callback.TriggerFire(IFACE_NAME, TrigInfo)
                End If

            End Select

          End If

        Next

      Catch pEx As Exception

      End Try

    Catch pEx As Exception

    End Try

  End Sub

#End Region

#Region "Action Properties"

  ''' <summary>
  ''' Defines the valid actions for this plug-in
  ''' </summary>
  ''' <remarks></remarks>
  Sub SetActions()
    Dim o As Object = Nothing
    If actions.Count = 0 Then
      actions.Add(o, "Set Channel")           ' 1
    End If
  End Sub

  ''' <summary>
  ''' Returns the action count
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Function ActionCount() As Integer
    SetActions()
    Return actions.Count
  End Function

  ''' <summary>
  ''' Returns the action name
  ''' </summary>
  ''' <param name="ActionNumber"></param>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  ReadOnly Property ActionName(ByVal ActionNumber As Integer) As String
    Get
      If Not ValidAction(ActionNumber) Then
        Return ""
      Else
        Return String.Format("{0}: {1}", IFACE_NAME, actions.Keys(ActionNumber - 1))
      End If
    End Get
  End Property

  ''' <summary>
  ''' Determines if an action is valid
  ''' </summary>
  ''' <param name="ActionIn"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Friend Function ValidAction(ByVal ActionIn As Integer) As Boolean
    SetActions()
    If ActionIn > 0 AndAlso ActionIn <= actions.Count Then
      Return True
    End If
    Return False
  End Function

#End Region

#Region "Action Interface"

  ''' <summary>
  ''' Builds the Action UI for display on the HomeSeer events page
  ''' </summary>
  ''' <param name="sUnique"></param>
  ''' <param name="ActInfo"></param>
  ''' <returns></returns>
  ''' <remarks>This function is called from the HomeSeer event page when an event is in edit mode.</remarks>
  Public Function ActionBuildUI(ByVal sUnique As String, ByVal ActInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As String

    Dim stb As New StringBuilder

    Try
      Dim UID As String = ActInfo.UID.ToString

      Dim action As New action
      If Not (ActInfo.DataIn Is Nothing) Then
        DeSerializeObject(ActInfo.DataIn, action)
      End If

      stb.AppendLine("<table cellspacing='0'>")

      Select Case ActInfo.TANumber
        Case Else

      End Select

      stb.AppendLine("</table>")

    Catch pEx As Exception
      '
      ' Process program exception
      '
      Call ProcessError(pEx, "ActionBuildUI()")
    End Try

    Return stb.ToString

  End Function

  ''' <summary>
  ''' When a user edits your event actions in the HomeSeer events, this function is called to process the selections.
  ''' </summary>
  ''' <param name="PostData"></param>
  ''' <param name="ActInfo"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function ActionProcessPostUI(ByVal PostData As Collections.Specialized.NameValueCollection,
                                      ByVal ActInfo As IPlugInAPI.strTrigActInfo) As IPlugInAPI.strMultiReturn

    Dim Ret As New HomeSeerAPI.IPlugInAPI.strMultiReturn

    Dim UID As Integer = ActInfo.UID
    Dim TANumber As Integer = ActInfo.TANumber

    ' When plug-in calls such as ...BuildUI, ...ProcessPostUI, or ...FormatUI are called and there is
    ' feedback or an error condition that needs to be reported back to the user, this string field 
    ' can contain the message to be displayed to the user in HomeSeer UI.  This field is cleared by
    ' HomeSeer after it is displayed to the user.
    Ret.sResult = ""

    ' We cannot be passed info ByRef from HomeSeer, so turn right around and return this same value so that if we want, 
    '   we can exit here by returning 'Ret', all ready to go.  If in this procedure we need to change DataOut or TrigInfo,
    '   we can still do that.
    Ret.DataOut = ActInfo.DataIn
    Ret.TrigActInfo = ActInfo

    If PostData Is Nothing Then Return Ret
    If PostData.Count < 1 Then Return Ret

    '
    ' DeSerializeObject
    '
    Dim action As New action
    If Not (ActInfo.DataIn Is Nothing) Then
      DeSerializeObject(ActInfo.DataIn, action)
    End If
    action.uid = UID

    Dim parts As Collections.Specialized.NameValueCollection = PostData

    Try

      Select Case TANumber
        Case AudioActions.SomeAction
          Dim actionName As String = GetEnumName(AudioActions.SomeAction)

          For Each sKey As String In parts.Keys

            If sKey Is Nothing Then Continue For
            If String.IsNullOrEmpty(sKey.Trim) Then Continue For

            Select Case True
              Case InStr(sKey, actionName & "TivoDevice_" & UID) > 0
                Dim ActionValue As String = parts(sKey)
                action.Item("TivoDevice") = ActionValue

              Case InStr(sKey, actionName & "TivoChannel_" & UID) > 0
                Dim ActionValue As String = parts(sKey)
                action.Item("TivoChannel") = ActionValue

            End Select
          Next

      End Select

      ' The serialization data for the plug-in object cannot be 
      ' passed ByRef which means it can be passed only one-way through the interface to HomeSeer.
      ' If the plug-in receives DataIn, de-serializes it into an object, and then makes a change 
      ' to the object, this is where the object can be serialized again and passed back to HomeSeer
      ' for storage in the HomeSeer database.

      ' SerializeObject
      If Not SerializeObject(action, Ret.DataOut) Then
        Ret.sResult = IFACE_NAME & " Error, Serialization failed. Signal Action not added."
        Return Ret
      End If

    Catch ex As Exception
      Ret.sResult = "ERROR, Exception in Action UI of " & IFACE_NAME & ": " & ex.Message
      Return Ret
    End Try

    ' All OK
    Ret.sResult = ""
    Return Ret
  End Function

  ''' <summary>
  ''' Determines if our action is proplery configured
  ''' </summary>
  ''' <param name="ActInfo"></param>
  ''' <returns>Return TRUE if the given action is configured properly</returns>
  ''' <remarks>There may be times when a user can select invalid selections for the action and in this case you would return FALSE so HomeSeer will not allow the action to be saved.</remarks>
  Public Function ActionConfigured(ByVal ActInfo As IPlugInAPI.strTrigActInfo) As Boolean

    Dim Configured As Boolean = True
    Dim UID As String = ActInfo.UID.ToString

    Dim action As New action
    If Not (ActInfo.DataIn Is Nothing) Then
      DeSerializeObject(ActInfo.DataIn, action)
    End If

    Select Case ActInfo.TANumber
      Case AudioActions.SomeAction
        If action.Item("TivoDevice") = "" Then Configured = False
        If action.Item("TivoChannel") = "" Then Configured = False

    End Select

    Return Configured

  End Function

  ''' <summary>
  ''' After the action has been configured, this function is called in your plugin to display the configured action
  ''' </summary>
  ''' <param name="ActInfo"></param>
  ''' <returns>Return text that describes the given action.</returns>
  ''' <remarks></remarks>
  Public Function ActionFormatUI(ByVal ActInfo As IPlugInAPI.strTrigActInfo) As String
    Dim stb As New StringBuilder

    Dim UID As String = ActInfo.UID.ToString

    Dim action As New action
    If Not (ActInfo.DataIn Is Nothing) Then
      DeSerializeObject(ActInfo.DataIn, action)
    End If

    Select Case ActInfo.TANumber
      Case Else

    End Select

    Return stb.ToString
  End Function

  ''' <summary>
  ''' Handles the HomeSeer Event Action
  ''' </summary>
  ''' <param name="ActInfo"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function HandleAction(ByVal ActInfo As IPlugInAPI.strTrigActInfo) As Boolean

    Dim UID As String = ActInfo.UID.ToString

    Try

      Dim action As New action
      If Not (ActInfo.DataIn Is Nothing) Then
        DeSerializeObject(ActInfo.DataIn, action)
      Else
        Return False
      End If

      Select Case ActInfo.TANumber
        Case Else

      End Select

    Catch pEx As Exception
      '
      ' Process Program Exception
      '
      hs.WriteLog(IFACE_NAME, "Error executing action: " & pEx.Message)
    End Try

    Return True

  End Function

#End Region

#End Region

End Module

Public Enum AudioActions
  <Description("Some Action")>
  SomeAction = 1
End Enum

Public Enum AudioTriggers
  <Description("Some Trigger")>
  SomeTrigger = 1
End Enum
