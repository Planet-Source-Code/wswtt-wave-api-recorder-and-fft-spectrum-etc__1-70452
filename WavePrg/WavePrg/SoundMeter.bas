Attribute VB_Name = "SoundMeter"
Option Explicit

'PCM “ÙônônÓ^
Public Type PCMFORM '“‘œ¬ûÈ44byteµƒônÓ^
    wRiffFormatTag As String * 4 '1~4¥Ê∑≈µƒ «RIFF◊÷¥Æ
    wfdataSize As Long '5~8¥Ê∑≈µƒ «ŸY¡œÖ^âK¥Û–°
' NOTE : ŸY¡œÖ^âK¥Û–°=(ôn∞∏¥Û–°-8)
    wFormatTag As String * 4 '9~12¥Ê∑≈µƒ «WAVE◊÷¥Æ
    wFormatName  As String * 4 '13~16¥Ê∑≈µƒ «◊”Ö^âK◊RÑe√˚∑Q
                           wCsize As Long '17~20¥Ê∑≈µƒ «◊”Ö^âK¥Û–°
                           wWavefmt As Integer '21~22¥Ê∑≈µƒ «¬ïôn∏Ò Ω,0x0001±ÌPCM∏Ò Ω
    wChannels As Integer '23~24¥Ê∑≈µƒ «¬ïµ¿îµ
    wSamplesPerSec As Long '25~28¥Ê∑≈µƒ√ø√Î»°ò”îµ
    wBytePerSec As Long '29~32¥Ê∑≈µƒ «√ø√ÎŸY¡œ¡ø
' NOTE : √ø√ÎŸY¡œ¡ø=(¬ïµ¿îµ*Œª‘™îµ*√ø√Î»°ò”îµ/8)
    wBytePerSample As Integer '33~34¥Ê∑≈µƒ «◊”Ö^âKŒª‘™ΩM
' NOTE : ◊”Ö^âKŒª‘™ΩM=(Œª‘™îµ/8)
    wBitsPerSample As Integer '35~36¥Ê∑≈µƒ «»°ò”ŒªΩM‘™îµ
    wData As String * 4 '¥Ê∑≈µƒ «data◊÷¥Æ
    wDataSize As Long 'åçÎH¬ïôn¥Û–°
' NOTE : ﬂ@ÇÄ÷µûÈôn∞∏¥Û–°úp»•ônÓ^(44BYTE)··µƒ÷µ
End Type

Public Type bdData '”√ÅÌÉ¶¥Ê8Œª‘™Îp¬ïµ¿µƒŸY¡œ
    bData1 As Byte '◊Û¬ïµ¿
    bData2 As Byte '”“¬ïµ¿
End Type

Public Type idData '”√ÅÌÉ¶¥Ê16Œª‘™Îp¬ïµ¿µƒŸY¡œ
    iData1 As Integer '◊Û¬ïµ¿
    iData2 As Integer '”“¬ïµ¿
End Type

Public pHandle As PCMFORM           'ônÓ^
'Public pHandle As PCMFORM, CuHandle As PCMFORM         'ônÓ^

'ônÓ^
Public LongData() As Integer  'wave data                            , CuDataL() As Integer
'Public LongData() As Integer, CuDataS() As Integer 'wave data                            , CuDataL() As Integer
Public startX As Long, endX As Long, FFTArr() As Double, FFTsize As Long, Haming As Boolean

Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

'eeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeee

Dim rc As Long                      ' return code
Dim ok As Boolean                   ' boolean return code
Dim volume As Long                      ' volume value
Dim volHmem As Long                     ' handle to volume memory

Public audByteArray() As Byte ' AUDINPUTARRAY
Public audIntArray() As Integer  ' AUDINPUTARRAY

Dim posval As Integer
Dim tempval As Integer

Private Const CALLBACK_FUNCTION = &H30000
Public Const CALLBACK_WINDOW = &H10000
Public Const MMIO_READ = &H0
Public Const MMIO_FINDCHUNK = &H10
Public Const MMIO_FINDRIFF = &H20

Public Const MM_WOM_OPEN = &H3BB   '  955
Public Const MM_WOM_DONE = &H3BD  '  957
Public Const MM_WOM_CLOSE = &H3BC  '  956

Public Const MMSYSERR_NOERROR = 0
Public Const SEEK_CUR = 1
Public Const SEEK_END = 2
Public Const SEEK_SET = 0
Public Const TIME_BYTES = &H4

Private Const MM_WIM_DATA = &H3C0
Private Const WHDR_DONE = &H1         '  done bit
Private Const GMEM_FIXED = &H0         ' Global Memory Flag used by GlobalAlloc functin

Type WAVEHDR
' The WAVEHDR user-defined type defines the header used to identify a waveform-audio buffer.
   lpData As Long          ' Address of the waveform buffer.
   dwBufferLength As Long  ' Length, in bytes, of the buffer.
   dwBytesRecorded As Long ' When the header is used in input, this member specifies how much
                           ' data is in the buffer.

   dwUser As Long          ' User data.
   dwFlags As Long         ' Flags supplying information about the buffer. Set equal to zero.
   dwLoops As Long         ' Number of times to play the loop. Set equal to zero.
   lpNext As Long          ' Not used
   Reserved As Long        ' Not used
End Type

Type WAVEINCAPS
' The WAVEINCAPS user-defined variable describes the capabilities of a waveform-audio input
' device.
   wMid As Integer         ' Manufacturer identifier for the device driver for the
                           ' waveform-audio input device. Manufacturer identifiers
                           ' are defined in Manufacturer and Product Identifiers in
                           ' the Platform SDK product documentation.
   wPid As Integer         ' Product identifier for the waveform-audio input device.
                           ' Product identifiers are defined in Manufacturer and Product
                           ' Identifiers in the Platform SDK product documentation.
   vDriverVersion As Long  ' Version number of the device driver for the
                           ' waveform-audio input device. The high-order byte
                           ' is the major version number, and the low-order byte
                           ' is the minor version number.
   szPname As String * 32  ' Product name in a null-terminated string.
   dwFormats As Long       ' Standard formats that are supported. See the Platform
                           ' SDK product documentation for more information.
   wChannels As Integer    ' Number specifying whether the device supports
                           ' mono (1) or stereo (2) input.
End Type

Type WAVEFORMAT
' The WAVEFORMAT user-defined type describes the format of waveform-audio data. Only
' format information common to all waveform-audio data formats is included in this
' user-defined type.
   wFormatTag As Integer      ' Format type. Use the constant WAVE_FORMAT_PCM Waveform-audio data
                              ' to define the data as PCM.
   nChannels As Integer       ' Number of channels in the waveform-audio data. Mono data uses one
                              ' channel and stereo data uses two channels.
   nSamplesPerSec As Long     ' Sample rate, in samples per second.
   nAvgBytesPerSec As Long    ' Required average data transfer rate, in bytes per second. For
                              ' example, 16-bit stereo at 44.1 kHz has an average data rate of
                              ' 176,400 bytes per second (2 channels ó 2 bytes per sample per
                              ' channel ó 44,100 samples per second).
   nBlockAlign As Integer     ' Block alignment, in bytes. The block alignment is the minimum atomic unit of data. For PCM data, the block alignment is the number of bytes used by a single sample, including data for both channels if the data is stereo. For example, the block alignment for 16-bit stereo PCM is 4 bytes (2 channels ó 2 bytes per sample).
   wBitsPerSample As Integer  ' For buffer estimation
   cbSize As Integer          ' Block size of the data.
End Type

Type mmioinfo
        dwFlags As Long
        fccIOProc As Long
        pIOProc As Long
        wErrorRet As Long
        htask As Long
        cchBuffer As Long
        pchBuffer As String
        pchNext As String
        pchEndRead As String
        pchEndWrite As String
        lBufOffset As Long
        lDiskOffset As Long
        adwInfo(4) As Long
        dwReserved1 As Long
        dwReserved2 As Long
        hmmio As Long
End Type

Type MMCKINFO
    ckid As Long
    ckSize As Long
    fccType As Long
    dwDataOffset As Long
    dwFlags As Long
End Type

Type MMTIME
        wType As Long
        u As Long
        X As Long
End Type
'Type AUDINPUTARRAY
'    bytes(5000) As Byte
'End Type
Private Declare Function waveInOpen Lib "winmm.dll" (lphWaveIn As Long, ByVal uDeviceID As Long, lpFormat As WAVEFORMAT, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
Private Declare Function waveInPrepareHeader Lib "winmm.dll" (ByVal hWaveIn As Long, lpWaveInHdr As WAVEHDR, ByVal uSize As Long) As Long
Private Declare Function waveInReset Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
Private Declare Function waveInStart Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
Private Declare Function waveInStop Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
Private Declare Function waveInUnprepareHeader Lib "winmm.dll" (ByVal hWaveIn As Long, lpWaveInHdr As WAVEHDR, ByVal uSize As Long) As Long
Private Declare Function waveInClose Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
Private Declare Function waveInGetDevCaps Lib "winmm.dll" Alias "waveInGetDevCapsA" (ByVal uDeviceID As Long, lpCaps As WAVEINCAPS, ByVal uSize As Long) As Long
Private Declare Function waveInGetNumDevs Lib "winmm.dll" () As Long
Private Declare Function waveInGetErrorText Lib "winmm.dll" Alias "waveInGetErrorTextA" (ByVal err As Long, ByVal lpText As String, ByVal uSize As Long) As Long
Private Declare Function waveInAddBuffer Lib "winmm.dll" (ByVal hWaveIn As Long, lpWaveInHdr As WAVEHDR, ByVal uSize As Long) As Long

Private Const MAXPNAMELEN = 32

Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hmem As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hmem As Long) As Long

' variables for managing wave file
Public format As WAVEFORMAT, formatOUT As WAVEFORMAT
Dim hmmioOut As Long
Dim mmckinfoParentIn As MMCKINFO
Dim mmckinfoSubchunkIn As MMCKINFO
Dim hWaveOut As Long
Dim bufferIn As Long
Dim hmemL As Long
Dim outHdr As WAVEHDR
Public numSamples As Long
Public drawFrom As Long
Public drawTo As Long
Public fFileLoaded As Boolean
Public fPlaying As Boolean

Public Declare Function waveOutPause Lib "winmm.dll" (ByVal hWaveOut As Long) As Long

Declare Function waveOutOpen Lib "winmm.dll" (lphWaveIn As Long, ByVal uDeviceID As Long, lpFormat As WAVEFORMAT, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
Declare Function waveOutGetPosition Lib "winmm.dll" (ByVal hWaveOut As Long, lpInfo As MMTIME, ByVal uSize As Long) As Long
Declare Function waveOutPrepareHeader Lib "winmm.dll" (ByVal hWaveIn As Long, lpWaveInHdr As WAVEHDR, ByVal uSize As Long) As Long
Declare Function waveOutReset Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
Declare Function waveOutUnprepareHeader Lib "winmm.dll" (ByVal hWaveIn As Long, lpWaveInHdr As WAVEHDR, ByVal uSize As Long) As Long
Declare Function waveOutClose Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
Declare Function waveOutGetDevCaps Lib "winmm.dll" Alias "waveInGetDevCapsA" (ByVal uDeviceID As Long, lpCaps As WAVEINCAPS, ByVal uSize As Long) As Long
Declare Function waveOutGetNumDevs Lib "winmm.dll" () As Long
Declare Function waveOutGetErrorText Lib "winmm.dll" Alias "waveInGetErrorTextA" (ByVal err As Long, ByVal lpText As String, ByVal uSize As Long) As Long
Declare Function waveOutWrite Lib "winmm.dll" (ByVal hWaveOut As Long, lpWaveOutHdr As WAVEHDR, ByVal uSize As Long) As Long
Declare Function mmioClose Lib "winmm.dll" (ByVal hmmio As Long, ByVal uFlags As Long) As Long
Declare Function mmioDescend Lib "winmm.dll" (ByVal hmmio As Long, lpck As MMCKINFO, lpckParent As MMCKINFO, ByVal uFlags As Long) As Long
Declare Function mmioDescendParent Lib "winmm.dll" Alias "mmioDescend" (ByVal hmmio As Long, lpck As MMCKINFO, ByVal X As Long, ByVal uFlags As Long) As Long
Declare Function mmioOpen Lib "winmm.dll" Alias "mmioOpenA" (ByVal szFileName As String, lpmmioinfo As mmioinfo, ByVal dwOpenFlags As Long) As Long
Declare Function mmioRead Lib "winmm.dll" (ByVal hmmio As Long, ByVal pch As Long, ByVal cch As Long) As Long
Declare Function mmioReadString Lib "winmm.dll" Alias "mmioRead" (ByVal hmmio As Long, ByVal pch As String, ByVal cch As Long) As Long
Declare Function mmioSeek Lib "winmm.dll" (ByVal hmmio As Long, ByVal lOffset As Long, ByVal iOrigin As Long) As Long
Declare Function mmioStringToFOURCC Lib "winmm.dll" Alias "mmioStringToFOURCCA" (ByVal sz As String, ByVal uFlags As Long) As Long
Declare Function mmioAscend Lib "winmm.dll" (ByVal hmmio As Long, lpck As MMCKINFO, ByVal uFlags As Long) As Long

Declare Sub CopyStructFromPtr Lib "kernel32" Alias "RtlMoveMemory" (struct As Any, ByVal ptr As Long, ByVal cb As Long)
Declare Sub CopyPtrFromStruct Lib "kernel32" Alias "RtlMoveMemory" (ByVal ptr As Long, struct As Any, ByVal cb As Long)
Declare Sub CopyStructFromString Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, ByVal Source As String, ByVal cb As Long)
Declare Function PostWavMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef hdr As WAVEHDR) As Long
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function ObjPtr Lib "MSVBVM60" Alias "VarPtr" (var As Object) As Long
Private Declare Function VarPtr Lib "MSVBVM60" (var As Any) As Long
   'Global Memory Flags
   'Global Const GMEM_FIXED = &H0
   'Global Const GMEM_MOVEABLE = &H2
   'Global Const GMEM_NOCOMPACT = &H10
   'Global Const GMEM_NODISCARD = &H20
   'Global Const GMEM_ZEROINIT = &H40
   'Global Const GMEM_MODIFY = &H80
   'Global Const GMEM_DISCARDABLE = &H100
   'Global Const GMEM_NOT_BANKED = &H1000
   'Global Const GMEM_SHARE = &H2000
   'Global Const GMEM_DDESHARE = &H2000
   'Global Const GMEM_NOTIFY = &H4000
   'Global Const GMEM_LOWER = GMEM_NOT_BANKED
   'Global Const GHND = (GMEM_MOVEABLE Or GMEM_ZEROINIT)
   'Global Const GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)

Private Msg As String * 200, hWaveIn As Long
Public NUM_BUFFERS As Integer
Public hmem() As Long, inHdr() As WAVEHDR
'Public hmem(NUM_BUFFERS - 1) As Long, inHdr(NUM_BUFFERS - 1) As WAVEHDR
Public BUFFER_SIZE As Long  '°°-------Œ¥”√£¨‘⁄ΩÁ√Ê÷– ‰»Î
Private Const DEVICEID = 0                 '   -1 ok too   , 1 bad   ??????????????????????/
Public Viewing As Boolean, Playing As Boolean, Recording As Boolean


                                                                                                     'µ§–ƒ∞Ê»®À˘”–VBƒ⁄«∂ASMº”øÏƒ⁄¥Ê ˝æ›∏¥÷∆
                                                                                                     'Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDest As Any, lpSource As Any, ByVal cBytes As Long)
                                                                                                     Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
                                                                                                     Private OpCode(200) As Byte, CodeStar As Long, opIndex As Long
                                                                                                     '*************************************************************************
                                                                                                     '**ƒ£ øÈ √˚£∫SuperCopyMemory VB∑∂¿˝
                                                                                                     '**Àµ       √˜£∫µ§–ƒ»Ìº˛‘⁄œﬂ…Ëº∆ ∞Ê»®À˘”–2007 - 2008(C)
                                                                                                     '**¥¥ Ω® »À£∫µ§–ƒ
                                                                                                     '**»’       ∆⁄£∫2007-09-03 22:13:43
                                                                                                     '**–ﬁ ∏ƒ »À£∫
                                                                                                     '**»’       ∆⁄£∫
                                                                                                     '**√Ë        ˆ£∫±»CopyMemoryªπ“™øÏµƒ∫Ø ˝SuperCopyMemory(),”¶”√‘⁄∏ﬂÀŸƒ⁄¥Ê∏¥÷∆–Ë«Û…œ
                                                                                                     '**∞Ê       ±æ£∫V1.0.0
                                                                                                     '**≤©øÕµÿ÷∑£∫http://hi.baidu.com/starwork/
                                                                                                     '**QQ     ∫≈¬Î£∫121877114
                                                                                                     '**E - mail:cnstarwork@126.com
                                                                                                     '*************************************************************************
                                                                                                     
                                                                                                     Public Sub SuperCopyMemory(ByVal lpDest As Long, ByVal lpSource As Long, ByVal cBytes As Long)
                                                                                                            CallWindowProc CodeStar, 0, lpDest, lpSource, cBytes
                                                                                                     End Sub
                                                                                                     
                                                                                                     Public Sub AsmIni()
                                                                                                            Dim i As Long
                                                                                                            CodeStar = (VarPtr(OpCode(0)) Or &HF) + 1
                                                                                                            opIndex = CodeStar - VarPtr(OpCode(0))
                                                                                                            For i = 0 To opIndex - 1
                                                                                                                OpCode(i) = &HCC
                                                                                                            Next
                                                                                                            AddByteToCode &H50: AddByteToCode &H53: AddByteToCode &H51: AddByteToCode &H56: AddByteToCode &H57: AddByteToCode &H8B
                                                                                                            AddByteToCode &H7C: AddByteToCode &H24: AddByteToCode 28: AddByteToCode &H8B: AddByteToCode &H74: AddByteToCode &H24
                                                                                                            AddByteToCode 32: AddByteToCode &H8B: AddByteToCode &H4C: AddByteToCode &H24: AddByteToCode 36: AddByteToCode &HB8
                                                                                                            i = 64
                                                                                                            AddLongToCode i: AddByteToCode &H8B: AddByteToCode &HD9: AddByteToCode &HFC: AddByteToCode &H3B: AddByteToCode &HC8
                                                                                                            AddByteToCode &H7C: AddByteToCode &H52: AddByteToCode &HC1: AddByteToCode &HE9: AddByteToCode &H6: AddByteToCode &HF
                                                                                                            AddByteToCode &H18: AddByteToCode &H46: AddByteToCode &H40: AddByteToCode &HF: AddByteToCode &H18: AddByteToCode &H47
                                                                                                            AddByteToCode &H40: AddByteToCode &HF: AddByteToCode &H6F: AddByteToCode &H6: AddByteToCode &HF: AddByteToCode &HE7: AddByteToCode &H7
                                                                                                            AddByteToCode &HF: AddByteToCode &H6F: AddByteToCode &H4E: AddByteToCode &H8: AddByteToCode &HF: AddByteToCode &HE7
                                                                                                            AddByteToCode &H4F: AddByteToCode &H8: AddByteToCode &HF: AddByteToCode &H6F: AddByteToCode &H56: AddByteToCode &H10
                                                                                                            AddByteToCode &HF: AddByteToCode &HE7: AddByteToCode &H57: AddByteToCode &H10: AddByteToCode &HF
                                                                                                            AddByteToCode &H6F: AddByteToCode &H5E: AddByteToCode &H18: AddByteToCode &HF: AddByteToCode &HE7: AddByteToCode &H5F
                                                                                                            AddByteToCode &H18: AddByteToCode &HF: AddByteToCode &H6F: AddByteToCode &H66: AddByteToCode &H20: AddByteToCode &HF
                                                                                                            AddByteToCode &HE7: AddByteToCode &H67: AddByteToCode &H20: AddByteToCode &HF: AddByteToCode &H6F
                                                                                                            AddByteToCode &H6E: AddByteToCode &H28: AddByteToCode &HF: AddByteToCode &HE7: AddByteToCode &H6F: AddByteToCode &H28
                                                                                                            AddByteToCode &HF: AddByteToCode &H6F: AddByteToCode &H76: AddByteToCode &H30: AddByteToCode &HF: AddByteToCode &HE7
                                                                                                            AddByteToCode &H77: AddByteToCode &H30: AddByteToCode &HF: AddByteToCode &H6F: AddByteToCode &H7E
                                                                                                            AddByteToCode &H38: AddByteToCode &HF: AddByteToCode &HE7: AddByteToCode &H7F: AddByteToCode &H38
                                                                                                            AddByteToCode &H3: AddByteToCode &HF0: AddByteToCode &H3: AddByteToCode &HF8: AddByteToCode &H49
                                                                                                            AddByteToCode &H75: AddByteToCode &HB3: AddByteToCode &HF
                                                                                                            AddByteToCode &H77: AddByteToCode &H8B: AddByteToCode &HCB: AddByteToCode &H48: AddByteToCode &H23
                                                                                                            AddByteToCode &HC8: AddByteToCode &H74: AddByteToCode &H2: AddByteToCode &HF3: AddByteToCode &HA4: AddByteToCode &H5F
                                                                                                            AddByteToCode &H5E: AddByteToCode &H59: AddByteToCode &H5B: AddByteToCode &H58: AddByteToCode &HC2
                                                                                                            AddByteToCode &H10: AddByteToCode &H0: AddByteToCode &HCC
                                                                                                     End Sub
                                                                                                     Public Sub AddByteToCode(bData As Byte)
                                                                                                            OpCode(opIndex) = bData
                                                                                                            opIndex = opIndex + 1
                                                                                                     End Sub
                                                                                                     Public Sub AddLongToCode(lData As Long)
                                                                                                            CopyMemory OpCode(opIndex), lData, 4
                                                                                                            opIndex = opIndex + 4
                                                                                                     End Sub
                                                                                                    'SuperCopyMemory VarPtr(dArr(0)), VarPtr(sArr(0)), 5000000
                                                                                                    'CopyMemory dArr(0), sArr(0), 5000000
                                                                                                    'SuperCopyMemory() Functionµ§–ƒ∞Ê»®À˘”–.◊™‘ÿ∑¢±Ì«Î◊¢√˜≥ˆ¥¶,≤¢Õ®÷™±æ»À



Private Sub waveInProc(ByVal hwi As Long, ByVal uMsg As Long, ByVal dwInstance As Long, ByRef hdr As WAVEHDR, ByVal dwParam2 As Long)
   If (uMsg = MM_WIM_DATA) Then
        If Viewing Then
              Dim nI As Long, Ya As Long, Xn As Long
              If hdr.dwUser = 8 Then
                    CopyMemory audByteArray(0), ByVal hdr.lpData, ByVal hdr.dwBufferLength
                                         DoEvents
                    'CopyStructFromPtr audbyteArray, hdr.lpData, hdr.dwBufferLength
                        If Recording Then
                                 Ya = UBound(LongData)
                                 Xn = hdr.dwBufferLength / (hdr.dwUser / 8)
                                 ReDim Preserve LongData(Ya + Xn)
                                 'CopyMemory LongData(Ya + 1), ByVal hdr.lpData, ByVal hdr.dwBufferLength
                                 For nI = Ya + 1 To Ya + Xn
                                         LongData(nI) = audByteArray(nI - Ya - 1)
                                         DoEvents
                                 'Debug.Print LongData(0), LongData(1)
                                 Next
                        End If
               ElseIf hdr.dwUser = 16 Then
                        CopyMemory audIntArray(0), ByVal hdr.lpData, ByVal hdr.dwBufferLength
                                 DoEvents
                        If Recording Then
                                 Ya = UBound(LongData)
                                 Xn = hdr.dwBufferLength / (hdr.dwUser / 8)
                                 ReDim Preserve LongData(Ya + Xn)
                                 CopyMemory LongData(Ya + 1), ByVal hdr.lpData, ByVal hdr.dwBufferLength
                                 'For nI = Ya + 1 To Ya + Xn
                                  '       LongData(nI) = audIntArray(nI - Ya - 1)
                                 'Debug.Print LongData(1), LongData(2)
                                 'Next
                                 DoEvents
                        End If
                        
               End If
               'Debug.Print audIntArray(6); audIntArray(16);audIntArray(36)
               'Debug.Print audByteArray(6); audByteArray(16); audByteArray(36)
        rc = waveInAddBuffer(hwi, hdr, Len(hdr))                  '             ??????????????????????????????????????
               
               'Debug.Print "ttttttttttttttttt"; hdr.lpData
                 'Debug.Print rc; Time; audbyteArray(3); audbyteArray(99)
         End If
   End If
End Sub
Public Function StartInput(BpS As Integer, Sps As Long, BUsampleNumber As Long) As Boolean 'BpS - wBitsPerSample , Sps - nSamplesPerSec, BUsampleNumber - BUFFER_Sample number
On Error GoTo err
        
   rc = waveInGetNumDevs
   If rc < 1 Then MsgBox "no soundcard.": Exit Function
   Dim ta As WAVEINCAPS
   rc = waveInGetDevCaps(0, ta, Len(ta))
    If rc <> 0 Then
        waveInGetErrorText rc, Msg, Len(Msg)
        MsgBox Msg
        StartInput = False
        Exit Function
    End If


    format.wFormatTag = 1
    format.nChannels = 1
    format.wBitsPerSample = BpS
    format.nSamplesPerSec = Sps
    format.nBlockAlign = format.nChannels * format.wBitsPerSample / 8
    format.nAvgBytesPerSec = format.nSamplesPerSec * format.nBlockAlign
    format.cbSize = 0
            ReDim audByteArray(BUsampleNumber - 1)
            ReDim audIntArray(BUsampleNumber - 1)
            Dim BUByteNumber As Long
    If format.wBitsPerSample = 8 Then
            BUByteNumber = BUsampleNumber     '◊™ªªŒ™ byte number Œ™¡ÀGlobalAlloc(&H40, BUsampleNumber)
    ElseIf format.wBitsPerSample = 16 Then
            BUByteNumber = BUsampleNumber * 2    '◊™ªªŒ™ byte number Œ™¡ÀGlobalAlloc(&H40, BUsampleNumber)
    Else
            MsgBox "BitsPerSample <> 8 or 16"
            Exit Function
    End If
    Dim i As Integer
    
        
    For i = 0 To NUM_BUFFERS - 1
        hmem(i) = GlobalAlloc(&H40, BUByteNumber)   '’‚¿Ô
        inHdr(i).lpData = GlobalLock(hmem(i))
        inHdr(i).dwBufferLength = BUByteNumber
        inHdr(i).dwFlags = 0
        inHdr(i).dwLoops = 0
        inHdr(i).dwUser = format.wBitsPerSample
    Next

    rc = waveInOpen(hWaveIn, DEVICEID, format, AddressOf waveInProc, 0, CALLBACK_FUNCTION)
    If rc <> 0 Then
        waveInGetErrorText rc, Msg, Len(Msg)
        MsgBox Msg
        StartInput = False
        Exit Function
    End If

    For i = 0 To NUM_BUFFERS - 1
        rc = waveInPrepareHeader(hWaveIn, inHdr(i), Len(inHdr(i)))
        If (rc <> 0) Then
            waveInGetErrorText rc, Msg, Len(Msg)
            MsgBox Msg
        End If
    Next

    For i = 0 To NUM_BUFFERS - 1
        rc = waveInAddBuffer(hWaveIn, inHdr(i), Len(inHdr(i)))
        If (rc <> 0) Then
            waveInGetErrorText rc, Msg, Len(Msg)
            MsgBox Msg
        End If
    Next
    Viewing = True
    rc = waveInStart(hWaveIn)
    StartInput = True
    Exit Function
err:
    StartInput = False
End Function
Public Function StopInput() As Integer
    On Error GoTo err
    Viewing = False
    rc = waveInReset(hWaveIn)
    'Debug.Print rc; 2
    rc = waveInStop(hWaveIn)
    'Debug.Print rc
    Dim i As Integer
    For i = 0 To NUM_BUFFERS - 1
        waveInUnprepareHeader hWaveIn, inHdr(i), Len(inHdr(i))
        GlobalFree hmem(i)
    Next
    waveInClose hWaveIn
    GlobalFree volHmem
    StopInput = 0
    Exit Function
err:
    StopInput = 1
End Function


'out   eeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeee
Sub waveOutProc(ByVal hwi As Long, ByVal uMsg As Long, ByVal dwInstance As Long, ByRef hdr As WAVEHDR, ByVal dwParam2 As Long)
' Wave IO Callback function
   'Debug.Print uMsg    ¥À”Ôæ‰”∞œÏcallback
   If (uMsg = MM_WOM_DONE) Then
      Playing = False
      'Debug.Print uMsg; "Playing = False  --- end"
   End If
End Sub

Sub StopPlay()
   'waveOutReset (hWaveOut)
   ' Close the waveout device
    rc = waveOutPause(hWaveOut)
    rc = waveOutReset(hWaveOut)
   ' Debug.Print rc; "waveOutReset"
    rc = waveOutClose(hWaveOut)
    'Debug.Print rc; "waveOutClose"
End Sub
Sub Play2(Channels As Integer, BitPerSample As Integer, SamplesPerSec As Long, lpCulongArr As Long, sizeINbyte As Long)
' Send audio buffer to wave output   ≤•∑≈ ˝◊È
    rc = waveOutGetNumDevs()
    Dim td As WAVEINCAPS
    rc = waveOutGetDevCaps(0, td, Len(td))
    If (rc <> 0) Then
        waveOutGetErrorText rc, Msg, Len(Msg)
        MsgBox Msg
        'Play = False
        Exit Sub
    End If
'waveFORMAT--------------formatout--- ' Get formatout info in Loadfile sub
'wFormatTag
'Format type. The following type is defined:    WAVE_FORMAT_PCM           Waveform-audio data is PCM.
'nChannels
'Number of channels in the waveform-audio data. Mono data uses one channel and stereo data uses two channels.
'nSamplesPerSec
'Sample rate, in samples per second.
'nAvgBytesPerSec
'Required average data transfer rate, in bytes per second. For example, 16-bit stereo at 44.1 kHz has an average data rate of 176,400 bytes per second (2 channels °™ 2 bytes per sample per channel °™ 44,100 samples per second).
'nBlockAlign
'Block alignment, in bytes. The block alignment is the minimum atomic unit of data. For PCM data, the block alignment is the number of bytes used by a single sample, including data for both channels if the data is stereo. For example, the block alignment for 16-bit stereo PCM is 4 bytes (2 channels °™ 2 bytes per sample).
    Dim i As Long
    i = SamplesPerSec * (BitPerSample / 8) * Channels
    
    formatOUT.cbSize = 0            'waveFORMAT--------------formatout
    formatOUT.nAvgBytesPerSec = i
    formatOUT.nBlockAlign = BitPerSample / 8
    formatOUT.nChannels = 1
    formatOUT.nSamplesPerSec = SamplesPerSec
    formatOUT.wBitsPerSample = BitPerSample
    formatOUT.wFormatTag = 1
    rc = waveOutOpen(hWaveOut, 0, formatOUT, AddressOf waveOutProc, 0, CALLBACK_FUNCTION)
    If (rc <> 0) Then
      GlobalFree (hmemL)
      waveOutGetErrorText rc, Msg, Len(Msg)
      MsgBox Msg
      Exit Sub
    End If

  ' GlobalFree hmemL
  ' hmemL = GlobalAlloc(&H40, 1024)
  ' bufferIn = GlobalLock(hmemL)
  
  
'WAVEHDR---------- outHdr
'lpData
'Address of the waveform buffer.
'dwBufferLength
'Length, in bytes, of the buffer.
'dwBytesRecorded
'When the header is used in input, this member specifies how much data is in the buffer.
'dwUser
'User data.
'dwFlags
'Flags supplying information about the buffer. The following values are defined:
    'WHDR_BEGINLOOP
    ''This buffer is the first buffer in a loop. This flag is used only with output buffers.
    'WHDR_DONE
    'Set by the device driver to indicate that it is finished with the buffer and is returning it to the application.
    'WHDR_ENDLOOP
    'This buffer is the last buffer in a loop. This flag is used only with output buffers.
    'WHDR_INQUEUE
    'Set by Windows to indicate that the buffer is queued for playback.
    'WHDR_PREPARED
    'Set by Windows to indicate that the buffer has been prepared with the waveInPrepareHeader or waveOutPrepareHeader function.
'dwLoops
'Number of times to play the loop. This member is used only with output buffers.
'wavehdr_tag
'Reserved.
'Reserved
'Reserved.
           ' Dim arrt() As Byte
           ' ReDim arrt(outHdr.dwBufferLength / 2 - 1)
           ' For I = 0 To outHdr.dwBufferLength / 2 - 1
           '         arrt(I) = Sin(I) * 30000
           ' Next

    outHdr.dwBufferLength = sizeINbyte
    outHdr.dwFlags = 0
    outHdr.dwLoops = 0
    outHdr.lpData = lpCulongArr
     
    'CopyMemory audByteArray(0), ByVal hdr.lpData, ByVal hdr.dwBufferLength
    'CopyMemory ByVal bufferIn, arrt(0), ByVal outHdr.dwBufferLength

    rc = waveOutPrepareHeader(hWaveOut, outHdr, Len(outHdr))
    If (rc <> 0) Then
      waveOutGetErrorText rc, Msg, Len(Msg)
      MsgBox Msg
    End If

    rc = waveOutWrite(hWaveOut, outHdr, Len(outHdr))
    If (rc <> 0) Then
      GlobalFree (hmemL)
    Else
      fPlaying = True                               '???
      Form1.Timer1.Enabled = True          '???
    End If
End Sub


   Sub SPLINE(X() As Double, Y() As Double, N As Long, YP1 As Double, YPN As Double, Y2() As Double) '»˝¥Œ—˘Ãı ∂˛Ω◊µº ˝
      Dim u() As Double, AAA As Double, BBB As Double, CCC As Double
      Dim P As Double, SIG As Double, QN As Double, UN As Double
      Dim i As Long, K As Long
      ReDim u(N)
      If YP1 > 9.9E+29 Then
          Y2(1) = 0
          u(1) = 0
      Else
          Y2(1) = -0.5
          AAA = (Y(2) - Y(1)) / (X(2) - X(1))
          u(1) = (3# / (X(2) - X(1))) * (AAA - YP1)
      End If
      For i = 2 To N - 1
          SIG = (X(i) - X(i - 1)) / (X(i + 1) - X(i - 1))
          P = SIG * Y2(i - 1) + 2#
          Y2(i) = (SIG - 1#) / P
          AAA = (Y(i + 1) - Y(i)) / (X(i + 1) - X(i))
          BBB = (Y(i) - Y(i - 1)) / (X(i) - X(i - 1))
          CCC = X(i + 1) - X(i - 1)
          u(i) = (6# * (AAA - BBB) / CCC - SIG * u(i - 1)) / P
      Next i
      If YPN > 9.9E+29 Then
          QN = 0#
          UN = 0#
      Else
          QN = 0.5
          AAA = YPN - (Y(N) - Y(N - 1)) / (X(N) - X(N - 1))
          UN = (3# / (X(N) - X(N - 1))) * AAA
      End If
      Y2(N) = (UN - QN * u(N - 1)) / (QN * Y2(N - 1) + 1#)
      For K = N - 1 To 1 Step -1
          Y2(K) = Y2(K) * Y2(K + 1) + u(K)
      Next K
   End Sub
Sub SPLINT(XA() As Double, Ya() As Double, Y2A() As Double, N As Long, X As Double, Y As Double)  '»˝¥Œ—˘Ãı
    Dim KLO As Double, KHI As Double, A  As Double, B  As Double, AAA As Double, BBB As Double, H As Double, K As Double

    KLO = 1
    KHI = N
1:   If KHI - KLO > 1 Then
        K = (KHI + KLO) / 2
        If XA(K) > X Then
            KHI = K
        Else
            KLO = K
        End If
        GoTo 1
    End If
    H = XA(KHI) - XA(KLO)
    If H = 0 Then
        MsgBox "  PAUSE  'BAD  XA  INPUT'"
        Exit Sub
    End If
    A = (XA(KHI) - X) / H
    B = (X - XA(KLO)) / H
    AAA = A * Ya(KLO) + B * Ya(KHI)
    BBB = (A ^ 3 - A) * Y2A(KLO) + (B ^ 3 - B) * Y2A(KHI)
    Y = AAA + BBB * (H ^ 2) / 6#
End Sub
Sub REALFT(ByRef DATA() As Double, N As Long, ISIGN As Long)                  ' µ∏µ¿˚“∂±‰ªª
      Dim THETA As Double, C1 As Double, C2 As Double, WPR As Double, WPI As Double, WR As Double, WI As Double
      Dim N2P3 As Long, i As Long, I1 As Long, I2 As Long, I3 As Long, I4 As Long
      Dim WRS As Single, WIS As Single, H1R As Double, H1I   As Double, H2R  As Double, H2I  As Double, WTEMP As Double
    THETA = 6.28318530717959 / 2# / N

    C1 = 0.5
    If ISIGN = 1 Then
        C2 = -0.5
        Call FOUR1(DATA(), N, 1)
    Else
        C2 = 0.5
        THETA = -THETA
    End If
    WPR = -2# * Sin(0.5 * THETA) ^ 2
    WPI = Sin(THETA)
    WR = 1# + WPR
    WI = WPI
    N2P3 = 2 * N + 3
    For i = 2 To N / 2 + 1
        I1 = 2 * i - 1
        I2 = I1 + 1
        I3 = N2P3 - I2
        I4 = I3 + 1
        WRS = CSng(WR)
        WIS = CSng(WI)
        H1R = C1 * (DATA(I1) + DATA(I3))
        H1I = C1 * (DATA(I2) - DATA(I4))
        H2R = -C2 * (DATA(I2) + DATA(I4))
        H2I = C2 * (DATA(I1) - DATA(I3))
        DATA(I1) = H1R + WRS * H2R - WIS * H2I
        DATA(I2) = H1I + WRS * H2I + WIS * H2R
        DATA(I3) = H1R - WRS * H2R + WIS * H2I
        DATA(I4) = -H1I + WRS * H2I + WIS * H2R
        WTEMP = WR
        WR = WR * WPR - WI * WPI + WR
        WI = WI * WPR + WTEMP * WPI + WI
    Next i
    If ISIGN = 1 Then
        H1R = DATA(1)
        DATA(1) = H1R + DATA(2)
        DATA(2) = H1R - DATA(2)
    Else
        H1R = DATA(1)
        DATA(1) = C1 * (H1R + DATA(2))
        DATA(2) = C1 * (H1R - DATA(2))
        Call FOUR1(DATA(), N, -1)
    End If
End Sub
Sub FOUR1(ByRef DATA() As Double, NN As Long, ISIGN As Long) ' µ∏µ¿˚“∂±‰ªª
      
      Dim N As Long, M As Long, i As Long, J As Long, TEMPR As Double, WTEMP  As Double, TEMPI As Double, MMAX As Long
      Dim THETA As Double, ISTEP As Long, WPR As Double, WPI As Double, WR As Double, WI As Double
      N = 2 * NN
      J = 1
      For i = 1 To N Step 2
          If J > i Then
              TEMPR = DATA(J)
              TEMPI = DATA(J + 1)
              DATA(J) = DATA(i)
              DATA(J + 1) = DATA(i + 1)
              DATA(i) = TEMPR
              DATA(i + 1) = TEMPI
          End If
          M = N / 2
          While M >= 2 And J > M
              J = J - M
              M = M / 2
          Wend
          J = J + M
      Next i
      MMAX = 2
      While N > MMAX
          ISTEP = 2 * MMAX
          THETA = 6.28318530717959 / (ISIGN * MMAX)
          WPR = -2# * Sin(0.5 * THETA) ^ 2
          WPI = Sin(THETA)
          WR = 1#
          WI = 0#
          For M = 1 To MMAX Step 2
              For i = M To N Step ISTEP
                  J = i + MMAX
                  TEMPR = CSng(WR) * DATA(J) - CSng(WI) * DATA(J + 1)
                  TEMPI = CSng(WR) * DATA(J + 1) + CSng(WI) * DATA(J)
                  DATA(J) = DATA(i) - TEMPR
                  DATA(J + 1) = DATA(i + 1) - TEMPI
                  DATA(i) = DATA(i) + TEMPR
                  DATA(i + 1) = DATA(i + 1) + TEMPI
              Next i
              WTEMP = WR
              WR = WR * WPR - WI * WPI + WR
              WI = WI * WPR + WTEMP * WPI + WI
          Next M
          MMAX = ISTEP
      Wend
End Sub

