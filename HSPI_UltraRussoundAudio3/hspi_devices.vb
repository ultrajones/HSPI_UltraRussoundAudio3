Imports HomeSeerAPI
Imports Scheduler
Imports HSCF
Imports HSCF.Communication.ScsServices.Service
Imports System.Text.RegularExpressions

Module hspi_devices

  ''' <summary>
  ''' Subroutine to create an Audio Zone Device
  ''' </summary>
  ''' <param name="device_id"></param>
  ''' <param name="dv_addr"></param>
  ''' <param name="bUpdateOnly"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function CreateAudioZoneDevice(ByVal device_id As String,
                                        ByVal dv_addr As String,
                                        ByVal bUpdateOnly As Boolean) As String

    Try

      If gAudioDevices.Any(Function(s) s.DeviceId = device_id) = True Then
        Dim AudioDevice As hspi_audio_device = gAudioDevices.Find(Function(s) s.DeviceId = device_id)
        Return AudioDevice.CreateAudioZoneDevice(dv_addr, bUpdateOnly)
      End If

    Catch pEx As Exception

    End Try

    Return False

  End Function

  ''' <summary>
  ''' Subroutine to create an ST2 Smart Tuner Device
  ''' </summary>
  ''' <param name="device_id"></param>
  ''' <param name="dv_addr"></param>
  ''' <param name="bUpdateOnly"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function CreateTunerDevice(ByVal device_id As String,
                                        ByVal dv_addr As String,
                                        ByVal bUpdateOnly As Boolean) As String

    Try

      If gAudioDevices.Any(Function(s) s.DeviceId = device_id) = True Then
        Dim AudioDevice As hspi_audio_device = gAudioDevices.Find(Function(s) s.DeviceId = device_id)
        Return AudioDevice.CreateTunerDevice(dv_addr, bUpdateOnly)
      End If

    Catch pEx As Exception

    End Try

    Return False

  End Function

  ''' <summary>
  ''' Create the HomeSeer Root Device
  ''' </summary>
  ''' <param name="strRootId"></param>
  ''' <param name="strRootName"></param>
  ''' <param name="dv_ref_child"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function CreateRootDevice(ByVal strRootId As String,
                                   ByVal strRootName As String,
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
      If strRootId = "Plugin" Then
        dv_name = "UltraRussoundAudio3 Plugin"
        dv_addr = String.Format("{0}-Root", strRootName.Replace(" ", "-"))
        dv_type = dv_name
      Else
        dv_name = strRootName
        dv_addr = String.Format("{0}-Root", strRootId, strRootName.Replace(" ", "-"))
        dv_type = strRootName
      End If

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

        Call WriteMessage(String.Format("Creating new HomeSeer {0} root device.", dv_name), MessageType.Debug)

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
        dv.Location2(hs) = IIf(strRootId = "Plugin", "Plug-ins", dv_type)
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
  ''' Determines if HomeSeer Device Exists
  ''' </summary>
  ''' <param name="dv_addr"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function DeviceExists(ByVal dv_addr As String) As Boolean

    Try

      Return hs.DeviceExistsAddress(dv_addr, False) > 0

    Catch pEx As Exception
      Call ProcessError(pEx, "DeviceExists")
      Return False
    End Try

  End Function

  ''' <summary>
  ''' Locates device by device by ref
  ''' </summary>
  ''' <param name="strDeviceAddr"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function LocateDeviceByRef(ByVal strDeviceAddr As String) As Object

    Dim objDevice As Object
    Dim dev_ref As Long = 0

    Try

      Long.TryParse(strDeviceAddr, dev_ref)
      If dev_ref > 0 Then
        objDevice = hs.GetDeviceByRef(dev_ref)
        If Not objDevice Is Nothing Then
          Return objDevice
        End If
      End If

    Catch pEx As Exception
      '
      ' Process the program exception
      '
      Call ProcessError(pEx, "LocateDeviceByRef")
    End Try
    Return Nothing ' No device found

  End Function

  ''' <summary>
  ''' Locates device by device by code
  ''' </summary>
  ''' <param name="strDeviceAddr"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function LocateDeviceByCode(ByVal strDeviceAddr As String) As Object

    Dim objDevice As Object
    Dim dev_ref As Long = 0

    Try

      dev_ref = hs.DeviceExistsCode(strDeviceAddr)
      objDevice = hs.GetDeviceByRef(dev_ref)
      If Not objDevice Is Nothing Then
        Return objDevice
      End If

    Catch pEx As Exception
      '
      ' Process the program exception
      '
      Call ProcessError(pEx, "LocateDeviceByCode")
    End Try
    Return Nothing ' No device found

  End Function

  ''' <summary>
  ''' Locates device by device address
  ''' </summary>
  ''' <param name="strDeviceAddr"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function LocateDeviceByAddr(ByVal strDeviceAddr As String) As Object

    Dim objDevice As Object
    Dim dev_ref As Long = 0

    Try

      dev_ref = hs.DeviceExistsAddress(strDeviceAddr, False)
      objDevice = hs.GetDeviceByRef(dev_ref)
      If Not objDevice Is Nothing Then
        Return objDevice
      End If

    Catch pEx As Exception
      '
      ' Process the program exception
      '
      Call ProcessError(pEx, "LocateDeviceByAddr")
    End Try
    Return Nothing ' No device found

  End Function

  ''' <summary>
  ''' Sets the HomeSeer string and device values
  ''' </summary>
  ''' <param name="dv_addr"></param>
  ''' <param name="dv_value"></param>
  ''' <remarks></remarks>
  Public Sub SetDeviceValue(ByVal dv_addr As String,
                            ByVal dv_value As String)

    Try

      WriteMessage(String.Format("{0}->{1}", dv_addr, dv_value), MessageType.Debug)

      Dim dv_ref As Integer = hs.DeviceExistsAddress(dv_addr, False)
      Dim bDeviceExists As Boolean = dv_ref <> -1

      If bDeviceExists = True Then

        WriteMessage(String.Format("Device address {0} was found.", dv_addr), MessageType.Debug)

        If IsNumeric(dv_value) Then

          Dim dblDeviceValue As Double = Double.Parse(hs.DeviceValueEx(dv_ref))
          Dim dblSensorValue As Double = Double.Parse(dv_value)

          If dblDeviceValue <> dblSensorValue Then
            hs.SetDeviceValueByRef(dv_ref, dblSensorValue, True)
          End If

        End If

      Else
        WriteMessage(String.Format("Device address {0} cannot be found.", dv_addr), MessageType.Debug)
      End If

    Catch pEx As Exception
      '
      ' Process program exception
      '
      Call ProcessError(pEx, "SetDeviceValue()")

    End Try

  End Sub

  ''' <summary>
  ''' Sets the HomeSeer device string
  ''' </summary>
  ''' <param name="dv_addr"></param>
  ''' <param name="dv_string"></param>
  ''' <remarks></remarks>
  Public Sub SetDeviceString(ByVal dv_addr As String,
                             ByVal dv_string As String)

    Try

      WriteMessage(String.Format("{0}->{1}", dv_addr, dv_string), MessageType.Debug)

      Dim dv_ref As Integer = hs.DeviceExistsAddress(dv_addr, False)
      Dim bDeviceExists As Boolean = dv_ref <> -1

      WriteMessage(String.Format("Device address {0} was found.", dv_addr), MessageType.Debug)

      If bDeviceExists = True Then

        hs.SetDeviceString(dv_ref, dv_string, True)

      Else
        WriteMessage(String.Format("Device address {0} cannot be found.", dv_addr), MessageType.Warning)
      End If

    Catch pEx As Exception
      '
      ' Process program exception
      '
      Call ProcessError(pEx, "SetDeviceString()")

    End Try

  End Sub

End Module
