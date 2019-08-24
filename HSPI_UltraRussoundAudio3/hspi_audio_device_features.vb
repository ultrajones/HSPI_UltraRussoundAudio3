Public Class hspi_audio_device_features

  Private _keypad_events As Boolean = False
  Private _source_control_events As Boolean = False
  Private _set_zone_on_off As Boolean = False
  Private _set_all_zone_on As Boolean = False
  Private _set_all_zone_off As Boolean = False
  Private _zone_source As Boolean = False
  Private _zone_volume As Boolean = False
  Private _tone_bass As Boolean = False
  Private _tone_treble As Boolean = False
  Private _tone_loudness As Boolean = False
  Private _tone_balance As Boolean = False
  Private _turn_on_volume As Boolean = False
  Private _background_color As Boolean = False
  Private _do_not_disturb As Boolean = False
  Private _party_mode As Boolean = False
  Private _all_zone_info As Boolean = False
  Private _display_string As Boolean = False
  Private _display_message As Boolean = False

  Public Sub New(ByVal device_model As String)

    Select Case device_model
      Case "CAS44"

        Me.KeypadEvents = True
        Me.SourceControlEvents = False
        Me.SetZoneOnOff = True
        Me.SetAllZoneOn = False
        Me.SetZoneOnOff = True
        Me.ZoneSource = True
        Me.ZoneVolume = True
        Me.ToneBass = True
        Me.ToneTreble = True
        Me.ToneLoudness = True
        Me.ToneBalance = True
        Me.TurnOnVolume = True
        Me.BackgroundColor = False
        Me.DoNotDisturb = False
        Me.PartyMode = False
        Me.AllZoneInfo = True
        Me.DisplayString = True
        Me.DisplayMessage = True

      Case "CAA66"

        Me.KeypadEvents = True
        Me.SourceControlEvents = True
        Me.SetZoneOnOff = True
        Me.SetAllZoneOn = False
        Me.SetZoneOnOff = True
        Me.ZoneSource = True
        Me.ZoneVolume = True
        Me.ToneBass = True
        Me.ToneTreble = True
        Me.ToneLoudness = True
        Me.ToneBalance = True
        Me.TurnOnVolume = True
        Me.BackgroundColor = False
        Me.DoNotDisturb = False
        Me.PartyMode = False
        Me.AllZoneInfo = True
        Me.DisplayString = True
        Me.DisplayMessage = True

      Case "CAM6.6", "CAV6.6"

        Me.KeypadEvents = True
        Me.SourceControlEvents = True
        Me.SetZoneOnOff = True
        Me.SetAllZoneOn = True
        Me.SetZoneOnOff = True
        Me.ZoneSource = True
        Me.ZoneVolume = True
        Me.ToneBass = True
        Me.ToneTreble = True
        Me.ToneLoudness = True
        Me.ToneBalance = True
        Me.TurnOnVolume = True
        Me.BackgroundColor = True
        Me.DoNotDisturb = True
        Me.PartyMode = True
        Me.AllZoneInfo = True
        Me.DisplayString = True
        Me.DisplayMessage = True

      Case "ACA-E5"

        Me.KeypadEvents = True
        Me.SourceControlEvents = True
        Me.SetZoneOnOff = True
        Me.SetAllZoneOn = True
        Me.SetZoneOnOff = True
        Me.ZoneSource = True
        Me.ZoneVolume = True
        Me.ToneBass = True
        Me.ToneTreble = True
        Me.ToneLoudness = True
        Me.ToneBalance = True
        Me.TurnOnVolume = True
        Me.BackgroundColor = False
        Me.DoNotDisturb = True
        Me.PartyMode = True
        Me.AllZoneInfo = True
        Me.DisplayString = True
        Me.DisplayMessage = True

    End Select

  End Sub

  Public Property KeypadEvents() As Boolean
    Get
      Return Me._keypad_events
    End Get
    Set(value As Boolean)
      Me._keypad_events = value
    End Set
  End Property

  Public Property SourceControlEvents() As Boolean
    Get
      Return Me._source_control_events
    End Get
    Set(value As Boolean)
      Me._source_control_events = value
    End Set
  End Property

  Public Property SetZoneOnOff() As Boolean
    Get
      Return Me._set_zone_on_off
    End Get
    Set(value As Boolean)
      Me._set_zone_on_off = value
    End Set
  End Property

  Public Property SetAllZoneOn() As Boolean
    Get
      Return Me._set_all_zone_on
    End Get
    Set(value As Boolean)
      Me._set_all_zone_on = value
    End Set
  End Property

  Public Property SetAllZoneOff() As Boolean
    Get
      Return Me._set_all_zone_off
    End Get
    Set(value As Boolean)
      Me._set_all_zone_off = value
    End Set
  End Property

  Public Property ZoneSource() As Boolean
    Get
      Return Me._zone_source
    End Get
    Set(value As Boolean)
      Me._zone_source = value
    End Set
  End Property

  Public Property ZoneVolume() As Boolean
    Get
      Return Me._zone_volume
    End Get
    Set(value As Boolean)
      Me._zone_volume = value
    End Set
  End Property

  Public Property ToneBass() As Boolean
    Get
      Return Me._tone_bass
    End Get
    Set(value As Boolean)
      Me._tone_bass = value
    End Set
  End Property

  Public Property ToneTreble() As Boolean
    Get
      Return Me._tone_treble
    End Get
    Set(value As Boolean)
      Me._tone_treble = value
    End Set
  End Property

  Public Property ToneLoudness() As Boolean
    Get
      Return Me._tone_loudness
    End Get
    Set(value As Boolean)
      Me._tone_loudness = value
    End Set
  End Property

  Public Property ToneBalance() As Boolean
    Get
      Return Me._tone_balance
    End Get
    Set(value As Boolean)
      Me._tone_balance = value
    End Set
  End Property

  Public Property TurnOnVolume() As Boolean
    Get
      Return Me._turn_on_volume
    End Get
    Set(value As Boolean)
      Me._turn_on_volume = value
    End Set
  End Property

  Public Property BackgroundColor() As Boolean
    Get
      Return Me._background_color
    End Get
    Set(value As Boolean)
      Me._background_color = value
    End Set
  End Property

  Public Property DoNotDisturb() As Boolean
    Get
      Return Me._do_not_disturb
    End Get
    Set(value As Boolean)
      Me._do_not_disturb = value
    End Set
  End Property

  Public Property PartyMode() As Boolean
    Get
      Return Me._party_mode
    End Get
    Set(value As Boolean)
      Me._party_mode = value
    End Set
  End Property

  Public Property AllZoneInfo() As Boolean
    Get
      Return Me._all_zone_info
    End Get
    Set(value As Boolean)
      Me._all_zone_info = value
    End Set
  End Property

  Public Property DisplayString() As Boolean
    Get
      Return Me._display_string
    End Get
    Set(value As Boolean)
      Me._display_string = value
    End Set
  End Property

  Public Property DisplayMessage() As Boolean
    Get
      Return Me._display_message
    End Get
    Set(value As Boolean)
      Me._display_message = value
    End Set
  End Property

End Class
