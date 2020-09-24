Attribute VB_Name = "Module2"

'== ACM API Declarations ================================================
Public Declare Function acmStreamOpen Lib "MSACM32" (hAS As Long, ByVal hADrv As Long, wfxSrc As WAVEFORMATEX, wfxDst As WAVEFORMATEX, ByVal wFltr As Long, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal fdwOpen As Long) As Long
Public Declare Function acmStreamClose Lib "MSACM32" (ByVal hAS As Long, ByVal dwClose As Long) As Long
Public Declare Function acmStreamPrepareHeader Lib "MSACM32" (ByVal hAS As Long, hASHdr As ACMSTREAMHEADER, ByVal dwPrepare As Long) As Long
Public Declare Function acmStreamUnprepareHeader Lib "MSACM32" (ByVal hAS As Long, hASHdr As ACMSTREAMHEADER, ByVal dwUnPrepare As Long) As Long
Public Declare Function acmStreamConvert Lib "MSACM32" (ByVal hAS As Long, hASHdr As ACMSTREAMHEADER, ByVal dwConvert As Long) As Long
Public Declare Function acmStreamReset Lib "MSACM32" (ByVal hAS As Long, ByVal dwReset As Long) As Long
Public Declare Function acmStreamSize Lib "MSACM32" (ByVal hAS As Long, ByVal cbInput As Long, dwOutBytes As Long, ByVal dwSize As Long) As Long
'== MCI Wave API Declarations ================================================
Declare Function timeGetTime Lib "winmm.dll" () As Long
'Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal SoundData As Any, ByVal uFlags As Long) As Long
Declare Function waveInOpen Lib "winmm.dll" (lphWaveIn As Long, ByVal uDeviceID As Long, lpFormat As WAVEFORMATEX, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
Declare Function waveInPrepareHeader Lib "winmm.dll" (ByVal hWaveIn As Long, wH As WAVEHDR, ByVal uSize As Long) As Long
Declare Function waveInUnprepareHeader Lib "winmm.dll" (ByVal hWaveIn As Long, wH As WAVEHDR, ByVal uSize As Long) As Long
Declare Function waveInAddBuffer Lib "winmm.dll" (ByVal hWaveIn As Long, wH As WAVEHDR, ByVal uSize As Long) As Long
Declare Function waveInStart Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
Declare Function waveInStop Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
Declare Function waveInReset Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
Declare Function waveInClose Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
Declare Function waveOutOpen Lib "winmm.dll" (lphWaveOut As Long, ByVal uDeviceID As Long, lpFormat As WAVEFORMATEX, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
Declare Function waveOutPrepareHeader Lib "winmm.dll" (ByVal hWaveOut As Long, wH As WAVEHDR, ByVal uSize As Long) As Long
Declare Function waveOutUnprepareHeader Lib "winmm.dll" (ByVal hWaveOut As Long, wH As WAVEHDR, ByVal uSize As Long) As Long
Declare Function waveOutWrite Lib "winmm.dll" (ByVal hWaveOut As Long, wH As WAVEHDR, ByVal uSize As Long) As Long
Declare Function waveOutClose Lib "winmm.dll" (ByVal hWaveOut As Long) As Long
Declare Function waveOutReset Lib "winmm.dll" (ByVal hWaveOut As Long) As Long
Declare Function waveOutPause Lib "winmm.dll" (ByVal hWaveOut As Long) As Long
Declare Function waveOutRestart Lib "winmm.dll" (ByVal hWaveOut As Long) As Long
'== Global Memory Functions ==================================================
Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Declare Function GlobalFree Lib "kernel32" (ByVal hmem As Long) As Long
Declare Function GlobalLock Lib "kernel32" (ByVal hmem As Long) As Long
Declare Function GlobalUnlock Lib "kernel32" (ByVal hmem As Long) As Long
Declare Sub CopyPTRtoBYTES Lib "Kernel32.dll" Alias "RtlMoveMemory" (ByRef ByteDest As Byte, ByVal PtrSrc As Long, ByVal length As Long)
Declare Sub CopyBYTEStoPTR Lib "Kernel32.dll" Alias "RtlMoveMemory" (ByVal PtrDest As Long, ByRef ByteSrc As Byte, ByVal length As Long)
Public Const GMEM_MOVEABLE = &H2&
Public Const GMEM_SHARE = &H2000&
Public Const GMEM_ZEROINIT = &H40&
Public Const TIMESLICE = 0.2            ' Time Slicing 1/5 Second
'== ACM API Constants ================================================
Public Const ACMERR_BASE = 512
Public Const ACMERR_NOTPOSSIBLE = (ACMERR_BASE + 0)
Public Const ACMERR_BUSY = (ACMERR_BASE + 1)
Public Const ACMERR_UNPREPARED = (ACMERR_BASE + 2)
Public Const ACMERR_CANCELED = (ACMERR_BASE + 3)
' AcmStreamSize Flags...
Public Const ACM_STREAMSIZEF_SOURCE = &H0&
Public Const ACM_STREAMSIZEF_DESTINATION = &H1&
Public Const ACM_STREAMSIZEF_QUERYMASK = &HF&
' acmStreamConvert Flags...
Public Const ACM_STREAMCONVERTF_BLOCKALIGN = &H4&
Public Const ACM_STREAMCONVERTF_START = &H10&
Public Const ACM_STREAMCONVERTF_END = &H20&
' Done Bits For ACMSTREAMHEADER.fdwStatus
Public Const ACMSTREAMHEADER_STATUSF_DONE = &H10000
Public Const ACMSTREAMHEADER_STATUSF_PREPARED = &H20000
Public Const ACMSTREAMHEADER_STATUSF_INQUEUE = &H100000
' Done Bits For acmStreamOpen Formats
Public Const ACM_STREAMOPENF_QUERY = &H1&
Public Const ACM_STREAMOPENF_ASYNC = &H2&
Public Const ACM_STREAMOPENF_NONREALTIME = &H4&
' Application Constants...
Public Const NoOfRings = 1                  ' Number Of Times In/Out Bound Calls Ring...
Public Const phoneHungUp = 3                ' Hangup Status Icon...
Public Const phoneRingIng = 2               ' Ringing Status Icon...
Public Const phoneAnswered = 1              ' Answered Status Icon...
Public Const mikeNO = 6
Public Const mikeOFF = 7
Public Const mikeON = 8
Public Const speakNO = 9
Public Const speakOFF = 10
Public Const speakON = 11
Public Const RingInId = 101                 ' Ringing InBound Sound...
Public Const RingOutId = 102                ' Ringing OutBound Sound...
' Toolbar constants...
Public Const tbCALL = 2
Public Const tbHANGUP = 3
Public Const tbAUTOANSWER = 5
'== MCI Wave API Declarations ================================================
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal SoundData As Any, ByVal uFlags As Long) As Long
'== TCP Port Array Processing Const.s ===================================
Public Const MINTCP = 1                     ' Minimum index for tcpsocket control instance
Public Const MAXTCP = 32                    ' Maximum index for tcpsocket control instance
'Public Const VOICEPORT = 701                ' Voice chat Port To Listen On...



Public Const NULLPORTID = 0                 ' Null Port ID - A Port ID That Will Never Be Used...
' Public Constants
Public Const MAXEXTRABYTES = 3          ' Maximum (Extra Bytes + 1) In Non PCM Wave Formats...
Public Const MAXBUFFERS = 500           ' Maximum Wave Buffer Array Member
Public Const MINBUFFERS = 0             ' Minimum Wave Buffer Array Member
'== MCI Wave Constants ================================================
' Sound Quality
Public Const c8_0kHz = 8000             ' 8.0 khz
Public Const c11_025kHz = 11025         ' 11.025 khz
Public Const c22_05kHz = 22050          ' 22.05 khz
Public Const c44_1kHz = 44010           ' 44.1 khz
' Sound Format
Public Const WAVE_FORMAT_PCM = &H1                  ' Microsoft Windows PCM Wave Format
Public Const WAVE_FORMAT_ADPCM = &H11               ' ADPCM Wave Format
Public Const WAVE_FORMAT_IMA_ADPCM = &H11           ' IMA ADPCM Wave Format
Public Const WAVE_FORMAT_DVI_ADPCM = &H11           ' DVI ADPCM Wave Format
Public Const WAVE_FORMAT_DSPGROUP_TRUESPEECH = &H22 ' DSP Group Wave Format
Public Const WAVE_FORMAT_GSM610 = &H31              ' GSM610 Wave Format
Public Const WAVE_FORMAT_MSN_AUDIO = &H32           ' MSN Audio Wave Format
' PCM Wave Format Types
Public Const WAVE_FORMAT_1M08 = &H1     '  11.025 kHz, Mono,   8-bit
Public Const WAVE_FORMAT_1M16 = &H4     '  11.025 kHz, Mono,   16-bit
Public Const WAVE_FORMAT_1S08 = &H2     '  11.025 kHz, Stereo, 8-bit
Public Const WAVE_FORMAT_1S16 = &H8     '  11.025 kHz, Stereo, 16-bit
Public Const WAVE_FORMAT_2M08 = &H10    '  22.05  kHz, Mono,   8-bit
Public Const WAVE_FORMAT_2M16 = &H40    '  22.05  kHz, Mono,   16-bit
Public Const WAVE_FORMAT_2S08 = &H20    '  22.05  kHz, Stereo, 8-bit
Public Const WAVE_FORMAT_2S16 = &H80    '  22.05  kHz, Stereo, 16-bit
Public Const WAVE_FORMAT_4M08 = &H100   '  44.1   kHz, Mono,   8-bit
Public Const WAVE_FORMAT_4M16 = &H400   '  44.1   kHz, Mono,   16-bit
Public Const WAVE_FORMAT_4S08 = &H200   '  44.1   kHz, Stereo, 8-bit
Public Const WAVE_FORMAT_4S16 = &H800   '  44.1   kHz, Stereo, 16-bit
'== Wave...Open() Constants ===========================================
Public Const WAVE_FORMAT_QUERY = &H1&   ' Query wave format flag
Public Const WAVE_MAPPER = (-1)         ' Maps To First Available Sound Device
Public Const WAVE_ALLOWSYNC = &H2&      ' Asynchronous playback flag
Public Const CALLBACK_WINDOW = &H10000  ' dwCallback is a HWND
Public Const CALLBACK_NULL = &H0&       ' no callback
'== MCI WaveHeader Bit Values In dwflags ==============================
Public Const WHDR_DONE = &H1&           '[00001][01] done bit
Public Const WHDR_PREPARED = &H2&       '[00010][02] set if this header has been prepared
Public Const WHDR_BEGINLOOP = &H4&      '[00100][04] loop start block
Public Const WHDR_ENDLOOP = &H8&        '[01000][08] loop end block
Public Const WHDR_INQUEUE = &H10&       '[10000][16] reserved for driver
'== MCI MM Return Codes ===============================================
Public Const ERROR_SHARING_VIOLATION = 32
Public Const MMSYSERR_NOERROR = 0                          '  no error
Public Const MMSYSERR_BASE = 0
Public Const MMSYSERR_ERROR = (MMSYSERR_BASE + 1)          '  unspecified error
Public Const MMSYSERR_BADDEVICEID = (MMSYSERR_BASE + 2)    '  device ID out of range
Public Const MMSYSERR_NOTENABLED = (MMSYSERR_BASE + 3)     '  driver failed enable
Public Const MMSYSERR_ALLOCATED = (MMSYSERR_BASE + 4)      '  device already allocated
Public Const MMSYSERR_INVALHANDLE = (MMSYSERR_BASE + 5)    '  device handle is invalid
Public Const MMSYSERR_NODRIVER = (MMSYSERR_BASE + 6)       '  no device driver present
Public Const MMSYSERR_NOMEM = (MMSYSERR_BASE + 7)          '  memory allocation error
Public Const MMSYSERR_NOTSUPPORTED = (MMSYSERR_BASE + 8)   '  function isn't supported
Public Const MMSYSERR_BADERRNUM = (MMSYSERR_BASE + 9)      '  error value out of range
Public Const MMSYSERR_INVALFLAG = (MMSYSERR_BASE + 10)     '  invalid flag passed
Public Const MMSYSERR_INVALPARAM = (MMSYSERR_BASE + 11)    '  invalid parameter passed
Public Const MMSYSERR_LASTERROR = (MMSYSERR_BASE + 11)     '  last error in range
Public Const WAVERR_BASE = 32
Public Const WAVERR_BADFORMAT = (WAVERR_BASE + 0)          '  unsupported wave format
Public Const WAVERR_STILLPLAYING = (WAVERR_BASE + 1)       '  still something playing
Public Const WAVERR_UNPREPARED = (WAVERR_BASE + 2)         '  header not prepared
Public Const WAVERR_LASTERROR = (WAVERR_BASE + 3)          '  last error in range
Public Const WAVERR_SYNC = (WAVERR_BASE + 3)               '  device is synchronous
'== flag values for wFlags parameter ==================================
Public Const SND_SYNC = &H0                 '  play synchronously (default)
Public Const SND_ASYNC = &H1                '  play asynchronously
Public Const SND_NODEFAULT = &H2            '  don't use default sound
Public Const SND_MEMORY = &H4               '  lpszSoundName points to a memory file
Public Const SND_LOOP = &H8                 '  loop the sound until next sndPlaySound
Public Const SND_NOSTOP = &H10              '  don't stop any currently playing sound
'== ACM User Defined Datatypes ================================================
Type WAVEFILTER
    cbStruct      As Long
    dwFilterTag   As Long
    fdwFilter     As Long
    dwReserved(5) As Long
End Type
Type ACMSTREAMHEADER            ' [ACM STREAM HEADER TYPE]
    cbStruct As Long            ' Size of header in bytes
    dwStatus As Long            ' Conversion status buffer
    dwUser As Long              ' 32 bits of user data specified by application
    pbSrc As Long               ' Source data buffer pointer
    cbSrcLength As Long         ' Source data buffer size in bytes
    cbSrcLengthUsed As Long     ' Source data buffer size used in bytes
    dwSrcUser As Long           ' 32 bits of user data specified by application
    cbDst As Long               ' Dest data buffer pointer
    cbDstLength As Long         ' Dest data buffer size in bytes
    cbDstLengthUsed As Long     ' Dest data buffer size used in bytes
    dwDstUser As Long           ' 32 bits of user data specified by application
    dwReservedDriver(9) As Long ' Reserved and should not be used
End Type
'== MCI User Defined Data Types...=======================================
Type WAVEHDR
    lpData As Long              ' pointer to locked data buffer
    dwBufferLength As Long      ' length of data buffer
    dwBytesRecorded As Long     ' used for input only
    dwUser As Long              ' for client's use
    dwFlags As Long             ' assorted flags (see defines)
    dwLoops As Long             ' loop control counter
    wavehdr_tag As Long         ' reserved for driver
    reserved As Long            ' reserved for driver
    hData As Long               ' handle to locked data buffer
End Type
Type WAVEFORMATEX
    wFormatTag As Integer       ' format type
    nChannels As Integer        ' number of channels (i.e. mono, stereo, etc.)
    nSamplesPerSec As Long      ' sample rate
    nAvgBytesPerSec As Long     ' for buffer estimation
    nBlockAlign As Integer      ' block size of data
    wBitsPerSample As Integer   ' Bits Per Sample
    cbSize As Integer           ' Size Of (FACT CHUNCK)
    xBytes(MAXEXTRABYTES) As Byte ' (FACT CHUNCK)
End Type

