Imports System.Text
Imports System.Web
Imports Scheduler
Imports HomeSeerAPI
Imports System.Collections.Specialized
Imports System.Text.RegularExpressions

Public Class hspi_webconfig
  Inherits clsPageBuilder

  Public hspiref As HSPI

  Dim TimerEnabled As Boolean

  ''' <summary>
  ''' Initializes new webconfig
  ''' </summary>
  ''' <param name="pagename"></param>
  ''' <remarks></remarks>
  Public Sub New(ByVal pagename As String)
    MyBase.New(pagename)
  End Sub

#Region "Page Building"

  ''' <summary>
  ''' Web pages that use the clsPageBuilder class and registered with hs.RegisterLink and hs.RegisterConfigLink will then be called through this function. 
  ''' A complete page needs to be created and returned.
  ''' </summary>
  ''' <param name="pageName"></param>
  ''' <param name="user"></param>
  ''' <param name="userRights"></param>
  ''' <param name="queryString"></param>
  ''' <param name="instance"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetPagePlugin(ByVal pageName As String, ByVal user As String, ByVal userRights As Integer, ByVal queryString As String, instance As String) As String

    Try

      Dim stb As New StringBuilder

      '
      ' Called from the start of your page to reset all internal data structures in the clsPageBuilder class, such as menus.
      '
      Me.reset()

      '
      ' Determine if user is authorized to access the web page
      '
      Dim LoggedInUser As String = hs.WEBLoggedInUser()
      Dim USER_ROLES_AUTHORIZED As Integer = WEBUserRolesAuthorized()

      '
      ' Handle any queries like mode=something
      '
      Dim parts As Collections.Specialized.NameValueCollection = Nothing
      If (queryString <> "") Then
        parts = HttpUtility.ParseQueryString(queryString)
      End If

      Dim Header As New StringBuilder
      'Header.AppendLine("<link type=""text/css"" href=""/hspi_ultrarussoundaudio3/css/hspi_ultrarussoundaudio3.css"" rel=""stylesheet"" />")
      Header.AppendLine("<link type=""text/css"" href=""/hspi_ultrarussoundaudio3/css/jquery.dataTables.min.css"" rel=""stylesheet"" />")
      Header.AppendLine("<link type=""text/css"" href=""/hspi_ultrarussoundaudio3/css/editor.dataTables.min.css"" rel=""stylesheet"" />")
      Header.AppendLine("<link type=""text/css"" href=""/hspi_ultrarussoundaudio3/css/buttons.dataTables.min.css"" rel=""stylesheet"" />")
      Header.AppendLine("<link type=""text/css"" href=""/hspi_ultrarussoundaudio3/css/select.dataTables.min.css"" rel=""stylesheet"" />")

      Header.AppendLine("<script type=""text/javascript"" src=""/hspi_ultrarussoundaudio3/js/jquery.dataTables.min.js""></script>")
      Header.AppendLine("<script type=""text/javascript"" src=""/hspi_ultrarussoundaudio3/js/dataTables.editor.min.js""></script>")
      Header.AppendLine("<script type=""text/javascript"" src=""/hspi_ultrarussoundaudio3/js/dataTables.buttons.min.js""></script>")
      Header.AppendLine("<script type=""text/javascript"" src=""/hspi_ultrarussoundaudio3/js/dataTables.select.min.js""></script>")

      Header.AppendLine("<script type=""text/javascript"" src=""/hspi_ultrarussoundaudio3/js/hspi_ultrarussoundaudio3_utility.js""></script>")
      Header.AppendLine("<script type=""text/javascript"" src=""/hspi_ultrarussoundaudio3/js/hspi_ultrarussoundaudio3_controllers.js""></script>")
      Me.AddHeader(Header.ToString)

      Dim pageTile As String = String.Format("{0} {1}", pageName, instance).TrimEnd
      stb.Append(hs.GetPageHeader(pageName, pageTile, "", "", False, False))

      '
      ' Start the page plug-in document division
      '
      stb.Append(clsPageBuilder.DivStart("pluginpage", ""))

      '
      ' A message area for error messages from jquery ajax postback (optional, only needed if using AJAX calls to get data)
      '
      stb.Append(clsPageBuilder.DivStart("divErrorMessage", "class='errormessage'"))
      stb.Append(clsPageBuilder.DivEnd)

      Me.RefreshIntervalMilliSeconds = 3000
      stb.Append(Me.AddAjaxHandlerPost("id=timer", pageName))

      If WEBUserIsAuthorized(LoggedInUser, USER_ROLES_AUTHORIZED) = False Then
        '
        ' Current user not authorized
        '
        stb.Append(WebUserNotUnauthorized(LoggedInUser))
      Else
        '
        ' Specific page starts here
        '
        stb.Append(BuildContent)
      End If

      '
      ' End the page plug-in document division
      '
      stb.Append(clsPageBuilder.DivEnd)

      '
      ' Add the body html to the page
      '
      Me.AddBody(stb.ToString)

      '
      ' Return the full page
      '
      Return Me.BuildPage()

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "GetPagePlugin")
      Return "error - " & Err.Description
    End Try

  End Function

  ''' <summary>
  ''' Builds the HTML content
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Function BuildContent() As String

    Try

      Dim stb As New StringBuilder

      stb.AppendLine("<table border='0' cellpadding='0' cellspacing='0' width='1000'>")
      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td width='1000' align='center' style='color:#FF0000; font-size:14pt; height:30px;'><strong><div id='divMessage'>&nbsp;</div></strong></td>")
      stb.AppendLine(" </tr>")
      stb.AppendLine(" <tr>")
      stb.AppendFormat("  <td class='tablecell'>{0}</td>", BuildTabs())
      stb.AppendLine(" </tr>")
      stb.AppendLine("</table>")

      Return stb.ToString

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "BuildContent")
      Return "error - " & Err.Description
    End Try

  End Function

  ''' <summary>
  ''' Builds the jQuery Tabss
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function BuildTabs() As String

    Try

      Dim stb As New StringBuilder
      Dim tabs As clsJQuery.jqTabs = New clsJQuery.jqTabs("oTabs", Me.PageName)
      Dim tab As New clsJQuery.Tab

      tabs.postOnTabClick = True

      tab.tabTitle = "Status"
      tab.tabDIVID = "tabStatus"
      tab.tabContent = "<div id='divStatus'>" & BuildTabStatus() & "</div>"
      tabs.tabs.Add(tab)

      tab = New clsJQuery.Tab
      tab.tabTitle = "Options"
      tab.tabDIVID = "tabOptions"
      tab.tabContent = "<div id='divOptions'></div>"
      tabs.tabs.Add(tab)

      tab = New clsJQuery.Tab
      tab.tabTitle = "Audio Controllers"
      tab.tabDIVID = "tabAudioControllers"
      tab.tabContent = "<div id='divAudioControllers'></div>"
      tabs.tabs.Add(tab)

      tab = New clsJQuery.Tab
      tab.tabTitle = "Audio Zone Devices"
      tab.tabDIVID = "tabAudioZoneDevices"
      tab.tabContent = "<div id='divAudioZoneDevices'></div>"
      tabs.tabs.Add(tab)

      'tab = New clsJQuery.Tab
      'tab.tabTitle = "Audio Zone Inputs"
      'tab.tabDIVID = "tabAudioZoneInputs"
      'tab.tabContent = "<div id='divAudioZoneInputs'></div>"
      'tabs.tabs.Add(tab)

      Return tabs.Build

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "BuildTabs")
      Return "error - " & Err.Description
    End Try

  End Function

  ''' <summary>
  ''' Build the Status Tab
  ''' </summary>
  ''' <param name="Rebuilding"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Function BuildTabStatus(Optional ByVal Rebuilding As Boolean = False) As String

    Try

      Dim stb As New StringBuilder

      stb.AppendLine(clsPageBuilder.FormStart("frmStatus", "frmStatus", "Post"))

      stb.AppendLine("<div>")
      stb.AppendLine("<table>")

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td>")
      stb.AppendLine("   <fieldset>")
      stb.AppendLine("    <legend> Plug-In Status </legend>")
      stb.AppendLine("    <table style=""width: 100%"">")
      stb.AppendLine("     <tr>")
      stb.AppendLine("      <td style=""width: 20%""><strong>Name:</strong></td>")
      stb.AppendFormat("    <td style=""text-align: right"">{0}</td>", IFACE_NAME)
      stb.AppendLine("     </tr>")
      stb.AppendLine("     <tr>")
      stb.AppendLine("      <td style=""width: 20%""><strong>Status:</strong></td>")
      stb.AppendFormat("    <td style=""text-align: right"">{0}</td>", "OK")
      stb.AppendLine("     </tr>")
      stb.AppendLine("     <tr>")
      stb.AppendLine("      <td style=""width: 20%""><strong>Version:</strong></td>")
      stb.AppendFormat("    <td style=""text-align: right"">{0}</td>", HSPI.Version)
      stb.AppendLine("     </tr>")
      stb.AppendLine("    </table>")
      stb.AppendLine("   </fieldset>")
      stb.AppendLine("  </td>")
      stb.AppendLine(" </tr>")

      stb.AppendLine("</table>")
      stb.AppendLine("</div>")

      stb.AppendLine(clsPageBuilder.FormEnd())

      If Rebuilding Then Me.divToUpdate.Add("divStatus", stb.ToString)

      Return stb.ToString

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "BuildTabStatus")
      Return "error - " & Err.Description
    End Try

  End Function

  ''' <summary>
  ''' Build the Options Tab
  ''' </summary>
  ''' <param name="Rebuilding"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Function BuildTabOptions(Optional ByVal Rebuilding As Boolean = False) As String

    Try

      Dim stb As New StringBuilder

      stb.AppendLine("<table cellspacing='0' width='100%'>")

      stb.Append(clsPageBuilder.FormStart("frmOptions", "frmOptions", "Post"))

      '
      ' Web Page Access
      '
      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tableheader' colspan='2'>Web Page Access</td>")
      stb.AppendLine(" </tr>")

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecell' style=""width: 20%"">Authorized User Roles</td>")
      stb.AppendFormat("  <td class='tablecell'>{0}</td>{1}", BuildWebPageAccessCheckBoxes, vbCrLf)
      stb.AppendLine(" </tr>")

      '
      ' Application Options
      '
      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tableheader' colspan='2'>Application Options</td>")
      stb.AppendLine(" </tr>")

      '
      ' Application Logging Level
      '
      Dim selLogLevel As New clsJQuery.jqDropList("selLogLevel", Me.PageName, False)
      selLogLevel.id = "selLogLevel"
      selLogLevel.toolTip = "Specifies the plug-in logging level."

      Dim itemValues As Array = System.Enum.GetValues(GetType(LogLevel))
      Dim itemNames As Array = System.Enum.GetNames(GetType(LogLevel))

      For i As Integer = 0 To itemNames.Length - 1
        Dim itemSelected As Boolean = IIf(gLogLevel = itemValues(i), True, False)
        selLogLevel.AddItem(itemNames(i), itemValues(i), itemSelected)
      Next
      selLogLevel.autoPostBack = True

      stb.AppendLine(" <tr>")
      stb.AppendFormat("  <td class='tablecell'>{0}</td>{1}", "Logging&nbsp;Level", vbCrLf)
      stb.AppendFormat("  <td class='tablecell'>{0}</td>{1}", selLogLevel.Build, vbCrLf)
      stb.AppendLine(" </tr>")

      stb.AppendLine("</table")

      stb.Append(clsPageBuilder.FormEnd())

      If Rebuilding Then Me.divToUpdate.Add("divOptions", stb.ToString)

      Return stb.ToString

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "BuildTabOptions")
      Return "error - " & Err.Description
    End Try

  End Function

  ''' <summary>
  ''' Build the Audio Controllers Tab
  ''' </summary>
  ''' <param name="Rebuilding"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Function BuildTabAudioControllers(Optional ByVal Rebuilding As Boolean = False) As String

    Try

      Dim stb As New StringBuilder

      stb.Append(clsPageBuilder.FormStart("frmAudioControllers", "frmAudioControllers", "Post"))

      stb.AppendLine("<table width='100%' class='display cell-border' id='table_devices' cellspacing='0'>")

      '
      ' HA7Net Configuration
      '
      stb.AppendLine(" <thead>")
      stb.AppendLine("  <tr>")
      stb.AppendLine("   <th>Device Name</th>")
      stb.AppendLine("   <th>Device Serial</th>")
      stb.AppendLine("   <th>Device Make</th>")
      stb.AppendLine("   <th>Device Model Number</th>")
      stb.AppendLine("   <th>Connection Type</th>")
      stb.AppendLine("   <th>Connection Address</th>")
      stb.AppendLine("   <th>Zones Defined</th>")
      stb.AppendLine("   <th>ST2 Tuner</th>")
      stb.AppendLine("   <th>SMS3 Media</th>")
      stb.AppendLine("   <th>Action</th>")
      stb.AppendLine("  </tr>")
      stb.AppendLine(" </thead>")

      stb.AppendLine(" <tbody>")
      Dim MyDataTable As DataTable = hspi_plugin.GetAudioDevices()
      For Each row As DataRow In MyDataTable.Rows
        stb.AppendFormat("  <tr id='{0}'>{1}", row("device_id"), vbCrLf)
        stb.AppendFormat("   <td>{0}</td>{1}", row("device_name"), vbCrLf)
        stb.AppendFormat("   <td>{0}</td>{1}", row("device_serial"), vbCrLf)
        stb.AppendFormat("   <td>{0}</td>{1}", row("device_make"), vbCrLf)
        stb.AppendFormat("   <td>{0}</td>{1}", row("device_model"), vbCrLf)
        stb.AppendFormat("   <td>{0}</td>{1}", row("device_conn"), vbCrLf)
        stb.AppendFormat("   <td>{0}</td>{1}", row("device_addr"), vbCrLf)
        stb.AppendFormat("   <td>{0}</td>{1}", row("device_zones"), vbCrLf)
        stb.AppendFormat("   <td>{0}</td>{1}", row("device_tuner_src"), vbCrLf)
        stb.AppendFormat("   <td>{0}</td>{1}", row("device_media_mgr"), vbCrLf)
        stb.AppendFormat("   <td>{0}</td>{1}", "", vbCrLf)
        stb.AppendLine("  </tr>")
      Next
      stb.AppendLine(" </tbody>")

      stb.AppendLine("</table")

      Dim strInfo As String = "The plug-in supports either a serial connection or a network connection when using a device like the Global Cache iTach Flex."
      Dim strHint As String = "When you modify an Audio Controller Device, a restart of the plug-in may be required."
      Dim strWarn As String = "Do not delete an Audio Controller Device if you have any HomeSeer devices configured because the devices may become orphaned."

      stb.AppendLine(" <div>&nbsp;</div>")
      stb.AppendLine(" <p>")
      stb.AppendFormat("<img alt='Info' src='/images/hspi_ultrarussoundaudio3/ico_info.gif' width='16' height='16' border='0' />&nbsp;{0}<br/>", strInfo)
      stb.AppendFormat("<img alt='Hint' src='/images/hspi_ultrarussoundaudio3/ico_hint.gif' width='16' height='16' border='0' />&nbsp;{0}<br/>", strHint)
      stb.AppendFormat("<img alt='Hint' src='/images/hspi_ultrarussoundaudio3/ico_warn.gif' width='16' height='16' border='0' />&nbsp;{0}<br/>", strWarn)
      stb.AppendLine(" </p>")

      stb.Append(clsPageBuilder.FormEnd())

      If Rebuilding Then Me.divToUpdate.Add("divAudioControllers", stb.ToString)

      Return stb.ToString

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "BuildTabAudioControllers")
      Return "error - " & Err.Description
    End Try

  End Function

  ''' <summary>
  ''' Build Audio Zone Devices
  ''' </summary>
  ''' <param name="device_id"></param>
  ''' <param name="device_type"></param>
  ''' <param name="Rebuilding"></param>
  ''' <returns></returns>
  Function BuildTabAudioZoneDevices(ByVal device_id As String,
                                    ByVal device_type As String,
                                    ByVal device_zone As String,
                                    Optional ByVal Rebuilding As Boolean = False) As String

    Try

      Dim stb As New StringBuilder

      stb.Append(clsPageBuilder.FormStart("frmAudioZones", "frmAudioZones", "Post"))

      '
      ' Get the Audio Device List
      '
      Dim AudioDeviceList As SortedList = hspi_plugin.GetAudioDeviceList()

      Dim selAudioDevice As New clsJQuery.jqDropList("selAudioDevice", Me.PageName, True)
      selAudioDevice.id = "selAudioDevice"
      selAudioDevice.toolTip = "Select the Audio Device"

      SyncLock gAudioDeviceLock

        For Each _AudioDevice As hspi_audio_device In gAudioDevices
          Dim value As String = _AudioDevice.DeviceId
          Dim desc As String = String.Format("Audio Device {0} [{1}] {2} {3}", _AudioDevice.DeviceId, _AudioDevice.ConnectionAddr, _AudioDevice.DeviceMake, _AudioDevice.DeviceModel)
          If device_id.Length = 0 Then device_id = value
          selAudioDevice.AddItem(desc, value, value = device_id)
        Next

        '
        ' Get the AudioZone Devices
        '
        Dim AudioDevice As hspi_audio_device = gAudioDevices.Find(Function(s) s.DeviceId = device_id)

        If IsNothing(AudioDevice) Then
          stb.AppendLine("<p>No Audio Devices were found.  Please make sure you have configured the plug-In to communicate with your Audio Device.</p>")
        Else
          Dim AudioZoneKeyTypes() As String = AudioDevice.GetAudioZoneKeyTypes()
          Dim selAudioZoneKeyType As New clsJQuery.jqDropList("selAudioZoneKeyType", Me.PageName, True)
          selAudioZoneKeyType.id = "selAudioZoneKeyType"
          selAudioZoneKeyType.toolTip = "Select the Audio Zone Device Type"

          For Each AudioZoneKeyType As String In AudioZoneKeyTypes
            If device_type.Length = 0 Then device_type = AudioZoneKeyType
            selAudioZoneKeyType.AddItem(AudioZoneKeyType, AudioZoneKeyType, device_type = AudioZoneKeyType)
          Next

          Dim selAudioDeviceZone As New clsJQuery.jqDropList("selAudioDeviceZone", Me.PageName, True)
          selAudioDeviceZone.id = "selAudioDeviceZone"

          If Regex.IsMatch(device_type, "^ST2 Smart Tuner") = True Then
            selAudioDeviceZone.toolTip = "Select the Smart Tuner Number"

            selAudioDeviceZone.AddItem("All Tuners", "", device_zone.Length = 0)
            For tt As Integer = 1 To 2
              Dim zone_name As String = String.Format("Tuner{0}", tt.ToString)
              selAudioDeviceZone.AddItem(zone_name, tt.ToString, device_zone = tt.ToString)
            Next
          Else
            selAudioDeviceZone.toolTip = "Select the Audio Zone"

            selAudioDeviceZone.AddItem("All Zones", "", device_zone.Length = 0)
            For zz As Integer = 1 To AudioDevice.DeviceZones
              Dim zone_name As String = String.Format("Zone{0}", zz.ToString)
              selAudioDeviceZone.AddItem(zone_name, zz.ToString, device_zone = zz.ToString)
            Next
          End If

          stb.AppendLine(" <div>")
          stb.AppendFormat("<b>{0}:</b>&nbsp;{1}&nbsp;{2}&nbsp;{3}{4}", "Audio Device", selAudioDevice.Build, selAudioZoneKeyType.Build, selAudioDeviceZone.Build, vbCrLf)
          stb.AppendLine(" </div>")

          stb.AppendLine("<table cellspacing='0' width='100%'>")

          '
          ' Russound Audio Device
          '
          stb.AppendLine(" <tr>")
          stb.AppendLine("  <td class='tableheader' colspan='10'>Audio Zone Devices</td>")
          stb.AppendLine(" </tr>")

          For Each AudioZone As hspi_audio_zone In AudioDevice.AudioZones
            AudioZone.DeviceExists = DeviceExists(AudioZone.DeviceAddr)
          Next

          '
          ' Render output based on selected Device Type
          '
          If Regex.IsMatch(device_type, "^ST2 Smart Tuner") = True Then
            '
            ' Process a ST2 Smart Tuner
            '
            stb.AppendLine(" <tr>")
            stb.AppendLine("  <td class='tablecolumn'>Device</td>")
            stb.AppendLine("  <td class='tablecolumn'>Controller</td>")
            stb.AppendLine("  <td class='tablecolumn'>Tuner</td>")
            stb.AppendLine("  <td class='tablecolumn'>Name</td>")
            stb.AppendLine("  <td class='tablecolumn'>Type</td>")
            stb.AppendLine("  <td class='tablecolumn'>Value</td>")
            stb.AppendLine("  <td class='tablecolumn'>String</td>")
            stb.AppendLine("  <td class='tablecolumn'>Last Change</td>")
            stb.AppendLine("  <td class='tablecolumn'>HomeSeer Device <input type=""checkbox"" title=""Check/Uncheck All""onclick=""javascript:toggleAll('chkAddDevice', this.checked)""/></td>")
            stb.AppendLine(" </tr>")

            Dim AudioZoneKeys() As String = AudioDevice.GetAudioZoneKeyNames(device_type)
            For Each AudioZoneKeyName As String In AudioZoneKeys

              '
              ' Determine if user wants to filter on a specific tuner
              '
              Dim tuner_start As Integer = 1
              Dim tuner_end As Integer = 2
              If device_zone.Length > 0 AndAlso IsNumeric(device_zone) AndAlso device_zone <= AudioDevice.DeviceZones Then
                tuner_start = device_zone
                tuner_end = device_zone
              End If

              Dim cc As Byte = 0
              Dim controller As Integer = 0
              Dim tuner As Integer = 0

              For tt = 0 To 1 Step 1
                controller = cc + 1
                tuner = tt + 1

                If tuner >= tuner_start AndAlso tuner <= tuner_end Then

                  Dim audioZoneId As String = String.Format("{0}.{1}.{2}", AudioDevice.DeviceId, cc, tt)
                  Dim dv_addr As String = String.Format("Russound{0}-{1}", audioZoneId, AudioZoneKeyName)

                  Dim AudioZone As hspi_audio_zone = AudioDevice.AudioZones.Find(Function(s) s.DeviceAddr = dv_addr)
                  If Not IsNothing(AudioZone) Then

                    Dim strChkInputName As String = "chkAddDevice"
                    Dim strChkInputValue As String = dv_addr
                    Dim strChkDisabled As String = String.Empty
                    Dim strChecked As String = String.Empty

                    If AudioZone.DeviceExists = False Then
                      strChkDisabled = ""
                      strChecked = ""
                    Else
                      strChkDisabled = "disabled"
                      strChecked = "checked"
                    End If

                    Dim strHSDevice As String = String.Format("<input type='{0}' name='{1}' value='{2}' {3} {4}>", "checkbox", strChkInputName, strChkInputValue, strChecked, strChkDisabled)

                    stb.AppendLine(" <tr>")
                    stb.AppendFormat("<td class='{0}'>{1}</td>", "tablecell", device_id)
                    stb.AppendFormat("<td class='{0}'>{1}</td>", "tablecell", controller.ToString)
                    stb.AppendFormat("<td class='{0}'>{1}</td>", "tablecell", tuner.ToString)
                    stb.AppendFormat("<td class='{0}'>{1}</td>", "tablecell", AudioDevice.GetKeyFriendlyName(AudioZoneKeyName))
                    stb.AppendFormat("<td class='{0}'>{1}</td>", "tablecell", device_type)
                    stb.AppendFormat("<td class='{0}' align='right'>{1}</td>", "tablecell", "-")
                    stb.AppendFormat("<td class='{0}' align='right'>{1}</td>", "tablecell", "")
                    stb.AppendFormat("<td class='{0}'>{1}</td>", "tablecell", "")
                    stb.AppendFormat("<td class='{0}'>{1}</td>", "tablecell", strHSDevice)
                    stb.AppendLine(" </tr>")

                  End If

                End If

              Next

            Next

          Else
            '
            ' Process an Audio Zone
            '
            stb.AppendLine(" <tr>")
            stb.AppendLine("  <td class='tablecolumn'>Device</td>")
            stb.AppendLine("  <td class='tablecolumn'>Controller</td>")
            stb.AppendLine("  <td class='tablecolumn'>Zone</td>")
            stb.AppendLine("  <td class='tablecolumn'>Name</td>")
            stb.AppendLine("  <td class='tablecolumn'>Type</td>")
            stb.AppendLine("  <td class='tablecolumn'>Value</td>")
            stb.AppendLine("  <td class='tablecolumn'>String</td>")
            stb.AppendLine("  <td class='tablecolumn'>Last Change</td>")
            stb.AppendLine("  <td class='tablecolumn'>HomeSeer Device <input type=""checkbox"" title=""Check/Uncheck All""onclick=""javascript:toggleAll('chkAddDevice', this.checked)""/></td>")
            stb.AppendLine(" </tr>")

            Dim AudioZoneKeys() As String = AudioDevice.GetAudioZoneKeyNames(device_type)
            For Each AudioZoneKeyName As String In AudioZoneKeys

              '
              ' Determine if user wants to filter on a specific zone
              '
              Dim zone_start As Integer = 1
              Dim zone_end As Integer = AudioDevice.DeviceZones
              If device_zone.Length > 0 AndAlso IsNumeric(device_zone) AndAlso device_zone <= AudioDevice.DeviceZones Then
                zone_start = device_zone
                zone_end = device_zone
              End If

              Dim controller As Integer = 0
              Dim zone As Integer = 0
              For cc As Integer = 0 To 5 Step 1
                controller = cc + 1
                For zz As Integer = 0 To 5 Step 1
                  zone += 1

                  If AudioDevice.DeviceZones >= zone Then

                    If zone >= zone_start AndAlso zone <= zone_end Then

                      Dim audioZoneId As String = String.Format("{0}.{1}.{2}", AudioDevice.DeviceId, cc, zz)
                      Dim dv_addr As String = String.Format("Russound{0}-{1}", audioZoneId, AudioZoneKeyName)

                      Dim AudioZone As hspi_audio_zone = AudioDevice.AudioZones.Find(Function(s) s.DeviceAddr = dv_addr)
                      If Not IsNothing(AudioZone) Then

                        Dim strChkInputName As String = "chkAddDevice"
                        Dim strChkInputValue As String = dv_addr
                        Dim strChkDisabled As String = String.Empty
                        Dim strChecked As String = String.Empty

                        If AudioZone.DeviceExists = False Then
                          strChkDisabled = ""
                          strChecked = ""
                        Else
                          strChkDisabled = "disabled"
                          strChecked = "checked"
                        End If

                        Dim strHSDevice As String = String.Format("<input type='{0}' name='{1}' value='{2}' {3} {4}>", "checkbox", strChkInputName, strChkInputValue, strChecked, strChkDisabled)

                        stb.AppendLine(" <tr>")
                        stb.AppendFormat("<td class='{0}'>{1}</td>", "tablecell", device_id)
                        stb.AppendFormat("<td class='{0}'>{1}</td>", "tablecell", controller.ToString)
                        stb.AppendFormat("<td class='{0}'>{1}</td>", "tablecell", zone.ToString)
                        stb.AppendFormat("<td class='{0}'>{1}</td>", "tablecell", AudioDevice.GetKeyFriendlyName(AudioZoneKeyName))
                        stb.AppendFormat("<td class='{0}'>{1}</td>", "tablecell", device_type)
                        stb.AppendFormat("<td class='{0}' align='right'>{1}</td>", "tablecell", AudioZone.GetPropertyValue(AudioZoneKeyName))
                        stb.AppendFormat("<td class='{0}' align='right'>{1}</td>", "tablecell", "")
                        stb.AppendFormat("<td class='{0}'>{1}</td>", "tablecell", "")
                        stb.AppendFormat("<td class='{0}'>{1}</td>", "tablecell", strHSDevice)
                        stb.AppendLine(" </tr>")

                      End If

                    End If

                  End If

                Next

              Next

            Next

          End If

          stb.AppendLine("</table")

          Dim jqButton1 As New clsJQuery.jqButton("btnAddDevices", "Save", Me.PageName, True)
          stb.AppendLine("<div>")
          stb.AppendLine(jqButton1.Build())
          stb.AppendLine("</div>")

        End If

      End SyncLock

      stb.Append(clsPageBuilder.FormEnd())

      If Rebuilding Then Me.divToUpdate.Add("divAudioZoneDevices", stb.ToString)

      Return stb.ToString

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "BuildTabAudioZoneDevices")
      Return "error - " & Err.Description
    End Try

  End Function

  ''' <summary>
  ''' Build Audio Zone Inputs
  ''' </summary>
  ''' <param name="device_id"></param>
  ''' <param name="Rebuilding"></param>
  ''' <returns></returns>
  Function BuildTabAudioZoneInputs(Optional ByVal device_id As String = "",
                                   Optional ByVal Rebuilding As Boolean = False) As String

    Try

      Dim stb As New StringBuilder

      stb.Append(clsPageBuilder.FormStart("frmAudioZoneInputs", "frmAudioZoneInputs", "Post"))

      '
      ' Get the Audio Device List
      '
      Dim AudioDeviceList As SortedList = hspi_plugin.GetAudioDeviceList()

      Dim selAudioDevice As New clsJQuery.jqDropList("selAudioDeviceInputs", Me.PageName, True)
      selAudioDevice.id = "selAudioDeviceInputs"
      selAudioDevice.toolTip = "Select the Audio Device"

      For Each _AudioDevice As hspi_audio_device In gAudioDevices
        Dim value As String = _AudioDevice.DeviceId
        Dim desc As String = String.Format("Audio Device {0} [{1}] {2} {3}", _AudioDevice.DeviceId, _AudioDevice.ConnectionAddr, _AudioDevice.DeviceMake, _AudioDevice.DeviceModel)
        If device_id.Length = 0 Then device_id = value
        selAudioDevice.AddItem(desc, value, value = device_id)
      Next

      If selAudioDevice.items.Count = 0 Then
        stb.AppendLine("<p>No Audio Controllers were found.  Please make sure you have configured the plug-in to communicate with your Audio Controller.</p>")
      Else

        stb.AppendLine(" <div>")
        stb.AppendFormat("<b>{0}:</b>&nbsp;{1}&nbsp;{2}", "Audio Device", selAudioDevice.Build, vbCrLf)
        stb.AppendLine(" </div>")

        stb.AppendLine("<table cellspacing='0' width='100%'>")

        '
        ' Russound Audio Device
        '
        stb.AppendLine(" <tr>")
        stb.AppendLine("  <td class='tableheader' colspan='10'>Audio Zone Inputs</td>")
        stb.AppendLine(" </tr>")

        stb.AppendLine("</table")

      End If

      stb.Append(clsPageBuilder.FormEnd())

      If Rebuilding Then Me.divToUpdate.Add("divAudioZoneInputs", stb.ToString)

      Return stb.ToString

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "BuildTabAudioZoneInputs")
      Return "error - " & Err.Description
    End Try

  End Function

  ''' <summary>
  ''' Build the Web Page Access Checkbox List
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function BuildWebPageAccessCheckBoxes()

    Try

      Dim stb As New StringBuilder

      Dim USER_ROLES_AUTHORIZED As Integer = WEBUserRolesAuthorized()

      Dim cb1 As New clsJQuery.jqCheckBox("chkWebPageAccess_Guest", "Guest", Me.PageName, True, True)
      Dim cb2 As New clsJQuery.jqCheckBox("chkWebPageAccess_Admin", "Admin", Me.PageName, True, True)
      Dim cb3 As New clsJQuery.jqCheckBox("chkWebPageAccess_Normal", "Normal", Me.PageName, True, True)
      Dim cb4 As New clsJQuery.jqCheckBox("chkWebPageAccess_Local", "Local", Me.PageName, True, True)

      cb1.id = "WebPageAccess_Guest"
      cb1.checked = CBool(USER_ROLES_AUTHORIZED And USER_GUEST)

      cb2.id = "WebPageAccess_Admin"
      cb2.checked = CBool(USER_ROLES_AUTHORIZED And USER_ADMIN)
      cb2.enabled = False

      cb3.id = "WebPageAccess_Normal"
      cb3.checked = CBool(USER_ROLES_AUTHORIZED And USER_NORMAL)

      cb4.id = "WebPageAccess_Local"
      cb4.checked = CBool(USER_ROLES_AUTHORIZED And USER_LOCAL)

      stb.Append(clsPageBuilder.FormStart("frmWebPageAccess", "frmWebPageAccess", "Post"))

      stb.Append(cb1.Build())
      stb.Append(cb2.Build())
      stb.Append(cb3.Build())
      stb.Append(cb4.Build())

      stb.Append(clsPageBuilder.FormEnd())

      Return stb.ToString

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "BuildWebPageAccessCheckBoxes")
      Return "error - " & Err.Description
    End Try

  End Function

#End Region

#Region "Page Processing"

  ''' <summary>
  ''' Post a message to this web page
  ''' </summary>
  ''' <param name="sMessage"></param>
  ''' <remarks></remarks>
  Sub PostMessage(ByVal sMessage As String)

    Try

      Me.divToUpdate.Add("divMessage", sMessage)

      Me.pageCommands.Add("starttimer", "")

      TimerEnabled = True

    Catch pEx As Exception

    End Try

  End Sub

  ''' <summary>
  ''' When a user clicks on any controls on one of your web pages, this function is then called with the post data. You can then parse the data and process as needed.
  ''' </summary>
  ''' <param name="page">The name of the page as registered with hs.RegisterLink or hs.RegisterConfigLink</param>
  ''' <param name="data">The post data</param>
  ''' <param name="user">The name of logged in user</param>
  ''' <param name="userRights">The rights of logged in user</param>
  ''' <returns>Any serialized data that needs to be passed back to the web page, generated by the clsPageBuilder class</returns>
  ''' <remarks></remarks>
  Public Overrides Function postBackProc(page As String, data As String, user As String, userRights As Integer) As String

    Try

      WriteMessage("Entered postBackProc() function.", MessageType.Debug)

      Dim postData As NameValueCollection = HttpUtility.ParseQueryString(data)

      '
      ' Write debug to console
      '
      If gLogLevel >= MessageType.Debug Then
        For Each keyName As String In postData.AllKeys
          'Console.WriteLine(String.Format("{0}={1}", keyName, postData(keyName)))
        Next
      End If

      Select Case postData("editor_action")
        Case "device-edit"
          Dim fields() As String = {"device_name", "device_serial", "device_make", "device_model", "device_conn", "device_addr", "device_zones", "device_tuner_src"}

          For Each field As String In fields
            Dim key As String = String.Format("data[{0}]", field)
            If postData.AllKeys.Contains(key) Then
              If postData(key).Trim.Length = 0 Then
                Return DatatableFieldError(field, "This Is a required field.")
              End If
            Else
              Return DatatableError("Unable To modify Audio Device due To an unexpected Error.")
            End If
          Next

          Dim device_id As Integer = Integer.Parse(postData("id"))
          Dim device_name As String = postData("data[device_name]").Trim
          Dim device_serial As String = postData("data[device_serial]").Trim
          Dim device_make As String = postData("data[device_make]").Trim
          Dim device_model As String = postData("data[device_model]").Trim
          Dim device_conn As String = postData("data[device_conn]").Trim
          Dim device_addr As String = postData("data[device_addr]").Trim

          Dim device_zones As Integer = Integer.Parse(postData("data[device_zones]"))
          Dim device_tuner_src As Integer = Integer.Parse(postData("data[device_tuner_src]"))
          Dim device_media_mgr As Integer = -1  ' Integer.Parse(postData("data[device_media_mgr]"))

          '
          ' Update Audio Device
          '
          Dim bSuccess As Boolean = UpdateAudioDevice(device_id, device_name, device_serial, device_make, device_model, device_conn, device_addr, device_zones, device_tuner_src, device_media_mgr)
          If bSuccess = False Then
            Return DatatableError("Unable To modify Audio Device due To an unexpected Error.")
          Else
            BuildTabAudioControllers(True)
            Me.pageCommands.Add("executefunction", "reDraw()")

            Return DatatableRowDevice(device_id, device_name, device_serial, device_make, device_model, device_conn, device_addr, device_zones, device_tuner_src, device_media_mgr)
          End If

        Case "device-create"

          Dim fields() As String = {"device_name", "device_serial", "device_make", "device_model", "device_conn", "device_addr", "device_zones", "device_tuner_src"}

          For Each field As String In fields
            Dim key As String = String.Format("data[{0}]", field)
            If postData.AllKeys.Contains(key) Then
              If postData(key).Trim.Length = 0 Then
                Return DatatableFieldError(field, "This Is a required field.")
              End If
            Else
              Return DatatableError("Unable To add Audio Device due To an unexpected Error.")
            End If
          Next

          Dim device_name As String = postData("data[device_name]").Trim
          Dim device_serial As String = postData("data[device_serial]").Trim
          Dim device_make As String = postData("data[device_make]").Trim
          Dim device_model As String = postData("data[device_model]").Trim
          Dim device_conn As String = postData("data[device_conn]").Trim
          Dim device_addr As String = postData("data[device_addr]").Trim

          Dim device_zones As Integer = Integer.Parse(postData("data[device_zones]"))
          Dim device_tuner_src As Integer = Integer.Parse(postData("data[device_tuner_src]"))
          Dim device_media_mgr As Integer = -1  ' Integer.Parse(postData("data[device_media_mgr]"))

          '
          ' Update the 1-Wire Device
          '
          Dim device_id As Integer = InsertAudioDevice(device_name, device_serial, device_make, device_model, device_conn, device_addr, device_zones, device_tuner_src, device_media_mgr)
          If device_id = False Then

            Return DatatableError("Unable To insert Audio Device device due To an Error.")

          Else
            BuildTabAudioControllers(True)
            Me.pageCommands.Add("executefunction", "reDraw()")

            Return DatatableRowDevice(device_id, device_name, device_serial, device_make, device_model, device_conn, device_addr, device_zones, device_tuner_src, device_media_mgr)
          End If

        Case "device-remove"
          Dim device_id As String = Val(postData("id[]"))

          Dim bSuccess As Boolean = DeleteAudioDevice(device_id)
          If bSuccess = False Then
            Return DatatableError("Unable To delete Audio Device due To an Error.")
          Else
            BuildTabAudioControllers(True)
            Me.pageCommands.Add("executefunction", "reDraw()")

            Return "{ }"
          End If

      End Select

      '
      ' Process the post data
      '
      Select Case postData("id")
        Case "tabStatus"
          BuildTabStatus(True)

        Case "tabOptions"
          BuildTabOptions(True)

        Case "tabAudioControllers"
          BuildTabAudioControllers(True)
          Me.pageCommands.Add("executefunction", "reDraw()")

        Case "tabAudioZoneDevices", "selAudioZoneKeyType", "selAudioDevice", "selAudioDeviceZone"
          Dim device_id As String = postData("selAudioDevice") & ""
          Dim device_type As String = postData("selAudioZoneKeyType") & ""
          Dim device_zone As String = postData("selAudioDeviceZone") & ""
          BuildTabAudioZoneDevices(device_id, device_type, device_zone, True)

        Case "btnAddDevices"

          Dim device_id As String = postData("selAudioDevice") & ""
          Dim device_type As String = postData("selAudioZoneKeyType") & ""
          Dim device_zone As String = postData("selAudioDeviceZone") & ""

          Select Case device_type
            Case "ST2 Smart Tuner"
              If Not (postData("chkAddDevice") Is Nothing) Then
                For Each Item As Object In postData("chkAddDevice").Split(",")
                  Dim strResults As String = CreateTunerDevice(device_id, Item, False)

                  If strResults <> String.Empty Then
                    PostMessage("Failed To process selected devices due to an Error: " & strResults)
                  End If
                Next
              End If
            Case Else
              If Not (postData("chkAddDevice") Is Nothing) Then
                For Each Item As Object In postData("chkAddDevice").Split(",")
                  Dim strResults As String = CreateAudioZoneDevice(device_id, Item, False)

                  If strResults <> String.Empty Then
                    PostMessage("Failed To process selected devices due to an Error: " & strResults)
                  End If
                Next
              End If
          End Select

          BuildTabAudioZoneDevices(device_id, device_type, device_zone, True)
          PostMessage("Audio Zone Device configuration saved.")

        Case "tabAudioZoneInputs", "selAudioDeviceInputs"
          Dim device_id As String = postData("selAudioDeviceInputs") & ""
          BuildTabAudioZoneInputs(device_id, True)

        Case "selLogLevel"
          gLogLevel = Int32.Parse(postData("selLogLevel"))
          hs.SaveINISetting("Options", "LogLevel", gLogLevel.ToString, gINIFile)

          PostMessage("The application logging level has been updated.")

        Case "WebPageAccess_Guest"

          Dim AUTH_ROLES As Integer = WEBUserRolesAuthorized()
          If postData("chkWebPageAccess_Guest") = "checked" Then
            AUTH_ROLES = AUTH_ROLES Or USER_GUEST
          Else
            AUTH_ROLES = AUTH_ROLES Xor USER_GUEST
          End If
          hs.SaveINISetting("WEBUsers", "AuthorizedRoles", AUTH_ROLES.ToString, gINIFile)

        Case "WebPageAccess_Normal"

          Dim AUTH_ROLES As Integer = WEBUserRolesAuthorized()
          If postData("chkWebPageAccess_Normal") = "checked" Then
            AUTH_ROLES = AUTH_ROLES Or USER_NORMAL
          Else
            AUTH_ROLES = AUTH_ROLES Xor USER_NORMAL
          End If
          hs.SaveINISetting("WEBUsers", "AuthorizedRoles", AUTH_ROLES.ToString, gINIFile)

        Case "WebPageAccess_Local"

          Dim AUTH_ROLES As Integer = WEBUserRolesAuthorized()
          If postData("chkWebPageAccess_Local") = "checked" Then
            AUTH_ROLES = AUTH_ROLES Or USER_LOCAL
          Else
            AUTH_ROLES = AUTH_ROLES Xor USER_LOCAL
          End If
          hs.SaveINISetting("WEBUsers", "AuthorizedRoles", AUTH_ROLES.ToString, gINIFile)

        Case "timer" ' This stops the timer and clears the message
          If TimerEnabled Then 'this handles the initial timer post that occurs immediately upon enabling the timer.
            TimerEnabled = False
          Else
            Me.pageCommands.Add("stoptimer", "")
            Me.divToUpdate.Add("divMessage", "&nbsp;")
          End If

      End Select

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "postBackProc")
    End Try

    Return MyBase.postBackProc(page, data, user, userRights)

  End Function

  ''' <summary>
  ''' Returns the Datatable Row JSON
  ''' </summary>
  ''' <param name="device_id"></param>
  ''' <param name="device_name"></param>
  ''' <param name="device_serial"></param>
  ''' <param name="device_make"></param>
  ''' <param name="device_model"></param>
  ''' <param name="device_conn"></param>
  ''' <param name="device_addr"></param>
  ''' <returns></returns>
  Private Function DatatableRowDevice(ByVal device_id As String,
                                      ByVal device_name As String,
                                      ByVal device_serial As String,
                                      ByVal device_make As String,
                                      ByVal device_model As String,
                                      ByVal device_conn As String,
                                      ByVal device_addr As String,
                                      ByVal device_zones As String,
                                      ByVal device_tuner_src As String,
                                      ByVal device_media_mgr As String) As String

    Try

      Dim sb As New StringBuilder
      sb.AppendLine("{")
      sb.AppendLine(" ""row"": { ")

      sb.AppendFormat(" ""{0}"": {1}, ", "DT_RowId", device_id)
      sb.AppendFormat(" ""{0}"": ""{1}"", ", "device_name", device_name)
      sb.AppendFormat(" ""{0}"": ""{1}"", ", "device_serial", device_serial)
      sb.AppendFormat(" ""{0}"": ""{1}"", ", "device_make", device_make)
      sb.AppendFormat(" ""{0}"": ""{1}"", ", "device_model", device_model)
      sb.AppendFormat(" ""{0}"": ""{1}"", ", "device_conn", device_conn)
      sb.AppendFormat(" ""{0}"": ""{1}"", ", "device_addr", device_addr)
      sb.AppendFormat(" ""{0}"": ""{1}"", ", "device_zones", device_zones)
      sb.AppendFormat(" ""{0}"": ""{1}"", ", "device_tuner_src", device_tuner_src)
      sb.AppendFormat(" ""{0}"": ""{1}"" ", "device_media_mgr", device_media_mgr)

      sb.AppendLine(" }")
      sb.AppendLine("}")

      Return sb.ToString

    Catch pEx As Exception
      Return "{ }"
    End Try

  End Function

  ''' <summary>
  ''' Returns the Datatable Error JSON
  ''' </summary>
  ''' <param name="errorString"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function DatatableError(ByVal errorString As String) As String

    Try
      Return String.Format("{{ ""error"": ""{0}"" }}", errorString)
    Catch pEx As Exception
      Return String.Format("{{ ""error"": ""{0}"" }}", pEx.Message)
    End Try

  End Function

  ''' <summary>
  ''' Returns the Datatable Field Error JSON
  ''' </summary>
  ''' <param name="fieldName"></param>
  ''' <param name="fieldError"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function DatatableFieldError(fieldName As String, fieldError As String) As String

    Try
      Return String.Format("{{ ""fieldErrors"": [ {{""name"": ""{0}"",""status"": ""{1}""}} ] }}", fieldName, fieldError)
    Catch pEx As Exception
      Return String.Format("{{ ""fieldErrors"": [ {{""name"": ""{0}"",""status"": ""{1}""}} ] }}", fieldName, pEx.Message)
    End Try

  End Function

#End Region

#Region "HSPI - Web Authorization"

  ''' <summary>
  ''' Returns the HTML Not Authorized web page
  ''' </summary>
  ''' <param name="LoggedInUser"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function WebUserNotUnauthorized(LoggedInUser As String) As String

    Try

      Dim sb As New StringBuilder

      sb.AppendLine("<table border='0' cellpadding='2' cellspacing='2' width='575px'>")
      sb.AppendLine("  <tr>")
      sb.AppendLine("   <td nowrap>")
      sb.AppendLine("     <h4>The Web Page You Were Trying To Access Is Restricted To Authorized Users ONLY</h4>")
      sb.AppendLine("   </td>")
      sb.AppendLine("  </tr>")
      sb.AppendLine("  <tr>")
      sb.AppendLine("   <td>")
      sb.AppendLine("     <p>This page is displayed if the credentials passed to the web server do not match the ")
      sb.AppendLine("      credentials required to access this web page.</p>")
      sb.AppendFormat("     <p>If you know the <b>{0}</b> user should have access,", LoggedInUser)
      sb.AppendFormat("      then ask your <b>HomeSeer Administrator</b> to check the <b>{0}</b> plug-in options", IFACE_NAME)
      sb.AppendFormat("      page to make sure the roles assigned to the <b>{0}</b> user allow access to this", LoggedInUser)
      sb.AppendLine("        web page.</p>")
      sb.AppendLine("  </td>")
      sb.AppendLine(" </tr>")
      sb.AppendLine(" </table>")

      Return sb.ToString

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "WebUserNotUnauthorized")
      Return "error - " & Err.Description
    End Try

  End Function

#End Region

End Class