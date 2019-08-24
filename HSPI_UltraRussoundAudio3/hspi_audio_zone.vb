Imports System.Text.RegularExpressions

Public Class hspi_audio_zone

  Private _dd As Byte = &H00
  Private _cc As Byte = &H00
  Private _zz As Byte = &H00
  Private _power As Byte = &H00
  Private _source As Byte = &H00
  Private _volume As Byte = &H00
  Private _turnon_volume As Byte = &H00
  Private _bass As Byte = &H00
  Private _treble As Byte = &H00
  Private _loudness As Byte = &H00
  Private _balance As Byte = &H00
  Private _system_onstate As Byte = &H00
  Private _shared_source As Byte = &H00
  Private _party_mode As Byte = &H00
  Private _dnd As Byte = &H00
  Private _keypad_bg_color As Byte = &H00
  Private _dv_addr As String = String.Empty
  Private _dv_exists As Boolean = False

  Public Sub New(ByVal dd As Byte, ByVal cc As Byte, ByVal zz As Byte, ByVal dv_addr As String)

    _dd = dd
    _cc = cc
    _zz = zz
    _dv_addr = dv_addr

  End Sub

  Public ReadOnly Property DeviceId() As Byte
    Get
      Return _dd
    End Get
  End Property

  Public ReadOnly Property ControllerId() As Byte
    Get
      Return _cc
    End Get
  End Property

  Public ReadOnly Property ZoneId() As Byte
    Get
      Return _zz
    End Get
  End Property

  Public Property Power() As Byte
    Get
      Return _power
    End Get
    Set(value As Byte)
      _power = value
    End Set
  End Property

  Public Property Source() As Byte
    Get
      Return _source
    End Get
    Set(value As Byte)
      _source = value
    End Set
  End Property

  Public Property Volume() As Byte
    Get
      Return _volume
    End Get
    Set(value As Byte)
      _volume = value
    End Set
  End Property

  Public Property TurnOnVolume() As Byte
    Get
      Return _turnon_volume
    End Get
    Set(value As Byte)
      _turnon_volume = value
    End Set
  End Property

  Public Property Bass() As Byte
    Get
      Return _bass
    End Get
    Set(value As Byte)
      _bass = value
    End Set
  End Property

  Public Property Treble() As Byte
    Get
      Return _treble
    End Get
    Set(value As Byte)
      _treble = value
    End Set
  End Property

  Public Property Loudness() As Byte
    Get
      Return _loudness
    End Get
    Set(value As Byte)
      _loudness = value
    End Set
  End Property

  Public Property Balance() As Byte
    Get
      Return _balance
    End Get
    Set(value As Byte)
      _balance = value
    End Set
  End Property

  Public Property SystemOnState() As Byte
    Get
      Return _system_onstate
    End Get
    Set(value As Byte)
      _system_onstate = value
    End Set
  End Property

  Public Property SharedSource() As Byte
    Get
      Return _shared_source
    End Get
    Set(value As Byte)
      _shared_source = value
    End Set
  End Property

  Public Property PartyMode() As Byte
    Get
      Return _party_mode
    End Get
    Set(value As Byte)
      _party_mode = value
    End Set
  End Property

  Public Property DoNotDisturb() As Byte
    Get
      Return _dnd
    End Get
    Set(value As Byte)
      _dnd = value
    End Set
  End Property

  Public Property KeypadBGColor() As String
    Get
      Return _keypad_bg_color
    End Get
    Set(value As String)
      _keypad_bg_color = value
    End Set
  End Property

  Public Property DeviceAddr() As String
    Get
      Return _dv_addr
    End Get
    Set(value As String)
      _dv_addr = value
    End Set
  End Property

  Public Property DeviceExists As Boolean
    Get
      Return _dv_exists
    End Get
    Set(value As Boolean)
      _dv_exists = value
    End Set
  End Property

  ''' <summary>
  ''' Returns the Zone Value for the selected property
  ''' </summary>
  ''' <param name="[property]"></param>
  ''' <returns></returns>
  Function GetPropertyValue(ByVal [property] As String) As Byte

    Try

      Select Case [property]
        Case "audiozone-power" : Return Power()
        Case "audiozone-source" : Return Source()
        Case "audiozone-partymode" : Return PartyMode()
        Case "audiozone-dnd" : Return DoNotDisturb
        Case "audiozone-turnon-vol" : Return TurnOnVolume()
        Case "audiozone-volume" : Return Volume()
        Case "audiozone-bass" : Return Bass()
        Case "audiozone-treble" : Return Treble()
        Case "audiozone-balance" : Return Balance()
        Case "audiozone-system-onstate" : Return SystemOnState()
        Case "audiozone-shared-source" : Return SharedSource()
        Case "audiozone-keypad-bg-color" : Return KeypadBGColor()
      End Select

    Catch pEx As Exception

    End Try

    Return &H00

  End Function

  ''' <summary>
  ''' Returns the Zone Value for the selected property
  ''' </summary>
  ''' <param name="[property]"></param>
  ''' <returns></returns>
  Function SetPropertyValue(ByVal [property] As String, ByVal value As Byte)

    Try

      Select Case [property]
        Case "audiozone-power" : Power() = value
        Case "audiozone-source" : Source() = value
        Case "audiozone-partymode" : PartyMode() = value
        Case "audiozone-dnd" : DoNotDisturb() = value
        Case "audiozone-turnon-vol" : TurnOnVolume() = value
        Case "audiozone-volume" : Volume() = value
        Case "audiozone-bass" : Bass() = value
        Case "audiozone-treble" : Treble() = value
        Case "audiozone-balance" : Balance() = value
        Case "audiozone-system-onstate" : SystemOnState() = value
        Case "audiozone-shared-source" : SharedSource() = value
        Case "audiozone-keypad-bg-color" : KeypadBGColor() = value
      End Select

    Catch pEx As Exception

    End Try

    Return &H00

  End Function

End Class
