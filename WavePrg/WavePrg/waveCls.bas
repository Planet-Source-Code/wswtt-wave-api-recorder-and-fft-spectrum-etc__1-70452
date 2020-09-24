Attribute VB_Name = "waveCls"
Option Explicit
Public CuHandle As PCMFORM         '檔頭  描述说明CuDataS()
Public CuDataS() As Integer             'cu wave data  整数数组　　８位播放＼保存时都有转换

Public frameN As Integer                       ' 帧长    点数
Public frameStepN As Integer                ' 帧间步长    点数
Public frameMS As Integer                     ' 帧长    毫秒数
Public frameStepMS As Integer              ' 帧间步长    毫秒数

Public PreEmphasisYesNo As Integer             '预加重否  －0 是没有检查（缺省值），1 为已检查，和 2 为变灰
Public PreEmphasisK As Single                       '预加重系数
Public set0YesNo As Integer                         ' set0否   去除直流偏置否


Public Sub makeCuM(pHandleP As PCMFORM, LongDataP() As Integer, startXP As Long, endXP As Long, PreEmphasisYesNoP As Integer, PreEmphasisKP As Single, set0YesNoP As Integer)
        On Error Resume Next
        
        'Form1.Text8.Text               wSamplesPerSecP
        'Form1.Checkbefore.Value    PreEmphasisYesNo  －0 是没有检查（缺省值），1 为已检查，和 2 为变灰
        'Form1.TextPreEmphasis       PreEmphasisK
        'Form1.Checkset0.Value        set0YesNo
        
        PreEmphasisYesNo = PreEmphasisYesNoP             '预加重否
        PreEmphasisK = PreEmphasisKP                            '预加重系数
        set0YesNo = set0YesNoP                                    ' set0否   去除直流偏置否
        
        CuHandle = pHandleP
                
        'CuHandle.wRiffFormatTag
        'CuHandle.wFormatTag
        'CuHandle.wFormatName
        'CuHandle.wCsize
        'CuHandle.wWavefmt
        CuHandle.wChannels = 1
            'CuHandle.wSamplesPerSec = wSamplesPerSecP  'ccccccccccccccccccccccccccccccccccccccchang  慢放快放？？
        CuHandle.wBytePerSample = pHandleP.wBitsPerSample / 8 * CuHandle.wChannels
        CuHandle.wBytePerSec = CuHandle.wSamplesPerSec * CuHandle.wBytePerSample
        'CuHandle.wBitsPerSample
        'CuHandle.wData
        CuHandle.wDataSize = (Abs(startXP - endXP) + 1) * pHandleP.wBitsPerSample / 8
        CuHandle.wfdataSize = CuHandle.wDataSize + 36
        Dim i As Long, tyt As Long
        ReDim CuDataS(Abs(startXP - endXP))
        'ReDim CuDataL(UBound(LongDataP))
        
        tyt = endXP - startXP  '避免 startXP - endXP ＝0 死循环
        If tyt = 0 Then tyt = 1
        
        If endXP >= UBound(LongDataP) Then endXP = endXP - 1   '避免   LongDataP(i + 1) 下标溢出
        For i = startXP To endXP Step Sgn(tyt)   '预加重
             If PreEmphasisYesNo = 1 Then
                 If pHandleP.wBitsPerSample = 8 Then
                    CuDataS(Abs(i - startXP)) = (LongDataP(i + 1) - 128) - PreEmphasisK * (LongDataP(i) - 128) + 128
                 Else
                    CuDataS(Abs(i - startXP)) = LongDataP(i + 1) - PreEmphasisK * LongDataP(i)
                    
                    If err.Number = 6 Then   '溢出（Overflow）的错误
                        CuDataS(Abs(i - startXP)) = LongDataP(i + 1)
                    End If
                    
                 End If
             Else
                CuDataS(Abs(i - startXP)) = LongDataP(i)
             End If
        Next i
        If set0YesNo = 1 Then
                Dim set0HA As Double
                set0HA = 0
                For i = 0 To Abs(startXP - endXP)    'set 0
                        set0HA = set0HA + CuDataS(i)
                Next i
                set0HA = set0HA / (Abs(startXP - endXP) + 1)
                If pHandleP.wBitsPerSample = 8 Then set0HA = set0HA - 128
                For i = 0 To Abs(startXP - endXP)  'set 0
                        CuDataS(i) = CuDataS(i) - set0HA
                Next i
        End If

End Sub


Public Function ShortTimeAveScope(ByRef DataInt() As Integer, XX As Long, sizeShortTimeN As Long) As Single
        On Error Resume Next
        Dim i As Long, heT As Long
        heT = 0
        For i = 0 To sizeShortTimeN - 1
            heT = heT + Abs(DataInt(XX + i - sizeShortTimeN / 2) - 128)
        Next i
         If CuHandle.wBitsPerSample = 8 Then
            ShortTimeAveScope = (heT) / 128 / sizeShortTimeN  '* CuHandle.wSamplesPerSec
         ElseIf CuHandle.wBitsPerSample = 16 Then
            ShortTimeAveScope = heT / 32768 / sizeShortTimeN '* CuHandle.wSamplesPerSec
         Else
         End If
End Function
Public Function ShortTimeAveEnergy(ByRef DataInt() As Integer, XX As Long, sizeShortTimeN As Long) As Single
        On Error Resume Next
        Dim i As Long, heT As Single
        heT = 0
         If CuHandle.wBitsPerSample = 8 Then
                For i = 0 To sizeShortTimeN - 1
                    heT = heT + (DataInt(XX + i - sizeShortTimeN / 2) - 128) * (DataInt(XX + i - sizeShortTimeN / 2) - 128)
                Next i
                ShortTimeAveEnergy = (heT) / 16384 / sizeShortTimeN * 20 '* CuHandle.wSamplesPerSec  128*128
         ElseIf CuHandle.wBitsPerSample = 16 Then
                For i = 0 To sizeShortTimeN - 1
                    heT = heT + (DataInt(XX + i - sizeShortTimeN / 2)) * (DataInt(XX + i - sizeShortTimeN / 2))
                Next i
                ShortTimeAveEnergy = heT / 1073741824 / sizeShortTimeN * 20 '* CuHandle.wSamplesPerSec  32768*32768
         Else
         End If
         'Debug.Print ShortTimeAveEnergy
End Function
Public Function ShortTimeAvecross0(ByRef DataInt() As Integer, XX As Long, sizeShortTimeN As Long) As Single
        On Error GoTo ttterrt
        Dim i As Long, heT As Long, Sss As Integer, ttSgn As Integer, iorB As Integer
        
        If CuHandle.wBitsPerSample = 8 Then iorB = 128 Else iorB = 0
         
        heT = 0
        Sss = 1   'or  -1 ?????     no=0
        For i = 0 To sizeShortTimeN - 1
            ttSgn = Sgn(DataInt(XX + i - sizeShortTimeN / 2) - iorB)
            If ttSgn <> 0 Then Sss = ttSgn: Exit For
        Next i
        For i = i To sizeShortTimeN - 1
            If Sgn(DataInt(XX + i - sizeShortTimeN / 2) - iorB) = -Sss Then heT = heT + 1: Sss = -Sss
        Next i
        'Debug.Print heT
        ShortTimeAvecross0 = heT / sizeShortTimeN '* CuHandle.wSamplesPerSec
        
    Exit Function
ttterrt:
    ShortTimeAvecross0 = Null
End Function
Public Function ShortTimeAvecrossPer(ByRef DataInt() As Integer, XX As Long, sizeShortTimeN As Long, Per As Single) As Single
        'ShortTimeAvecross0  --->  ShortTimeAvecrossPer
   On Error GoTo ttterr
        Dim i As Long, heT As Long, axT As Long, Sss As Integer, iorB As Integer, PerD As Integer, D1 As Integer, D2 As Integer
        
        If CuHandle.wBitsPerSample = 8 Then
                iorB = 128
                PerD = 128 * Per
        ElseIf CuHandle.wBitsPerSample = 16 Then
                iorB = 0
                PerD = 32768 * Per
        End If
        D1 = iorB - PerD
        D2 = iorB + PerD
         
        heT = 0
        Sss = 0   'or  -1 ?????     no=0
        For i = 0 To sizeShortTimeN - 1
            axT = XX + i - sizeShortTimeN / 2
            If DataInt(axT) < D1 Then
                    Sss = -1
                    Exit For
            ElseIf DataInt(axT) > D2 Then
                    Sss = 1
                    Exit For
            End If
        Next i
        For i = i To sizeShortTimeN - 1
            axT = XX + i - sizeShortTimeN / 2
            If DataInt(axT) < D1 And Sss = 1 Then
                    heT = heT + 1
                    Sss = -1
            ElseIf DataInt(axT) > D2 And Sss = -1 Then
                    heT = heT + 1
                    Sss = 1
            End If
        Next i
        'Debug.Print heT

        ShortTimeAvecrossPer = heT / sizeShortTimeN '* CuHandle.wSamplesPerSec
    Exit Function
ttterr:
    ShortTimeAvecrossPer = Null
End Function

