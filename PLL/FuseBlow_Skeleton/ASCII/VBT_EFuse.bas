Attribute VB_Name = "VBT_EFuse"
Option Explicit
Public Enum eI2CEFuseRegs
    eI2CReg1E = &H1E
    eI2CReg1F = &H1F
    eI2CReg1D = &H1D
    X_Coordinate = &H7F00&
    Y_Coordinate = &H7F01&
    LOT_ID_0 = &H7F02&
    LOT_ID_1 = &H7F03&
    LOT_ID_2 = &H7F04&
    LOT_ID_3 = &H7F05&
    Wafer_ID = &H7F06&
    Manufacture_ID_1 = &H7F07&
    Manufacture_ID_2 = &H7F08&
    Prod_SN = &H7F09&
    Trim_Byte_0 = &H7F0A&
    Trim_Byte_1 = &H7F0B&
    PLL_TRIM_0 = &H7F0C&
    PLL_TRIM_1 = &H7F0D&
    OSC_TRIM = &H7F0E&
    ADC_TRIM = &H7F0F&
    DAC_TRIM = &H7F10&
    LDO_TRIM_0 = &H7F11&
    LDO_TRIM_1 = &H7F12&
    LDO_TRIM_2 = &H7F13&
    BUCK_TRIM_0 = &H7F14&
    BUCK_TRIM_1 = &H7F15&
    BUCK_TRIM_2 = &H7F16&
    BUBO_TRIM_0 = &H7F17&
    BUBO_TRIM_1 = &H7F18&
    Vtg_TRIM_1 = &H7F19&
    Vtg_TRIM_2 = &H7F1A&
    CPU_TRIM = &H7F1B&
End Enum

Public Vtg_Trimmed_Flag As New SiteBoolean
Public Freq_Trimmed_Flag As New SiteBoolean
Public slFreq_TrimCode As New SiteLong
Public slVoltage_TrimCode As New SiteLong

Public Function EFuseOTPErase()
    On Error GoTo errHandler
    Dim sldata As New SiteLong
    Dim ret_data As New SiteLong
    
    Dim lI As Long

    ' Load level and timing.
    'TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    'sldata = &H7&
    'Call I2CEFuseRead(OSC_TRIM, ret_data, sldata)
    
    sldata = &H8C&
    Call I2CEFuseWrite(&H1E&, sldata)
    'Call I2CEFuseRead(&H1E&, ret_data, sldata)
    'Debug.Print "0x1E=", sldata(0) 'print 0x1E

    sldata = &H9F&
    Call I2CEFuseWrite(&H1F&, sldata)
    'Call I2CEFuseRead(&H1F&, ret_data, sldata)
    'Debug.Print "0x1F=", sldata(0) 'print 0x1F

    sldata = &H7E&
    Call I2CEFuseWrite(&H1D&, sldata)
    'Call I2CEFuseRead(&H1D&, ret_data, sldata)
    'Debug.Print "0x1D=", sldata(0) 'print 0x1D
    
    'Reset all efuse as 0

    sldata = 0
    Call I2CEFuseWrite(&H7F00& + &HE&, sldata)
    Call I2CEFuseWrite(&H7F00& + &H19&, sldata)
    Call I2CEFuseWrite(&H7F00& + &H1A&, sldata)
    
    'enable efuse programming
    sldata = 0
    Call I2CEFuseWrite(&H1D&, sldata)
    
    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function EFuseOTPBurning()
    On Error GoTo errHandler

    Dim ADC As New SiteLong
    Dim ADC_Gain_CMPST As New SiteLong
    Dim ADC_Offset_CMPST As New SiteLong
    Dim BUBO_0 As New SiteLong
    Dim BUBO_1 As New SiteLong
    Dim BUCK_0 As New SiteLong
    Dim BUCK_1 As New SiteLong
    Dim BUCK_1_ISET_Code As New SiteLong
    Dim BUCK_1_VSEL_Code As New SiteLong
    Dim BUCK_2 As New SiteLong
    Dim BUCK_2_ISET_Code As New SiteLong
    Dim BUCK_2_VSEL_Code As New SiteLong
    Dim BUCK_3_ISET_Code As New SiteLong
    Dim BUCK_3_VSEL_Code As New SiteLong
    Dim BUCK_JUMP_SIZE_0 As New SiteLong
    Dim BUCK_JUMP_SIZE_1 As New SiteLong
    Dim CPU As New SiteLong
    Dim CPU_MAX_FREQ As New SiteLong
    Dim DAC As New SiteLong
    Dim DAC_Gain_CMPST As New SiteLong
    Dim DAC_Offset_CMPST As New SiteLong
    Dim Device_ID1 As New SiteLong
    Dim Device_ID2 As New SiteLong
    Dim Device_ID3 As New SiteLong
    Dim Device_ID4 As New SiteLong
    Dim Freq_Trim As New SiteLong
    Dim INT_OSC As New SiteLong
    Dim INT_OSC_TRIM_ As New SiteLong
    Dim LDO_0 As New SiteLong
    Dim LDO_1 As New SiteLong
    Dim LDO_1_VCMP_Code As New SiteLong
    Dim LDO_1_VOFST_Code As New SiteLong
    Dim LDO_1_VSEL_Code As New SiteLong
    Dim LDO_2 As New SiteLong
    Dim LDO_2_VCMP_Code As New SiteLong
    Dim LDO_2_VOFST_Code As New SiteLong
    Dim LDO_2_VSEL_Code As New SiteLong
    Dim LDO_3_VCMP_Code As New SiteLong
    Dim LDO_3_VOFST_Code As New SiteLong
    Dim LDO_3_VSEL_Code As New SiteLong
    Dim Manufacture_week As New SiteLong
    Dim Manufacture_Year As New SiteLong
    Dim OSC As New SiteLong
    Dim OSC_TRIM_0 As New SiteLong
    Dim OSC_TRIM_1 As New SiteLong
    Dim PLL_0 As New SiteLong
    Dim PLL_1 As New SiteLong
    Dim PLL_Div_0 As New SiteLong
    Dim PLL_Div_1 As New SiteLong
    Dim PLL_Div_2 As New SiteLong
    Dim PLL_Div_3 As New SiteLong
    Dim Production_Line_ID As New SiteLong
    Dim TSEN As New SiteLong
    Dim TSEN_CAL_FACTOR_TRIM As New SiteLong
    Dim TSEN_EN As New SiteLong
    Dim TSEN_OFST As New SiteLong
    Dim Volt_Trim As New SiteLong
    Dim Wafer_Size As New SiteLong
    Dim Wafer_Type As New SiteLong
    Dim WF_Good As New SiteLong
    Dim X_Wafer_Coordinate As New SiteLong
    Dim Y_Wafer_Coordinate As New SiteLong
    Dim sldata As New SiteLong
    Dim slReadback(27) As New SiteLong
    Dim slReadBackExpect(27) As New SiteLong
    
    Dim slReadData As New SiteLong
    
    Dim lI As Long

    ' Load level and timing.
    thehdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    
    ADC = 1
    ADC_Gain_CMPST = &H7
    ADC_Offset_CMPST = &H1
    BUBO_0 = 1
    BUBO_1 = 1
    BUCK_0 = 1
    BUCK_1 = 1
    BUCK_1_ISET_Code = &H18
    BUCK_1_VSEL_Code = &H7
    BUCK_2 = 1
    BUCK_2_ISET_Code = 0
    BUCK_2_VSEL_Code = 0
    BUCK_3_ISET_Code = &H0
    BUCK_3_VSEL_Code = &H6
    BUCK_JUMP_SIZE_0 = &HA
    BUCK_JUMP_SIZE_1 = &H3
    CPU = 1
    CPU_MAX_FREQ = &HFC
    DAC = 1
    DAC_Gain_CMPST = &H8
    DAC_Offset_CMPST = &H0
    Device_ID1 = &HB6
    Device_ID2 = &H5C
    Device_ID3 = &H37
    Device_ID4 = &HF1
    Freq_Trim = &HB
    INT_OSC = 1
    INT_OSC_TRIM_ = &H21
    LDO_0 = 1
    LDO_1 = 1
    LDO_1_VCMP_Code = &H7
    LDO_1_VOFST_Code = &H2
    LDO_1_VSEL_Code = &H3
    LDO_2 = 1
    LDO_2_VCMP_Code = &H0
    LDO_2_VOFST_Code = &H1
    LDO_2_VSEL_Code = &H3
    LDO_3_VCMP_Code = &H5
    LDO_3_VOFST_Code = &H1
    LDO_3_VSEL_Code = &H2
    Manufacture_week = &H23
    Manufacture_Year = &H7D
    OSC = 1
    OSC_TRIM_0 = &H4
    OSC_TRIM_1 = &H0
    PLL_0 = 1
    PLL_1 = 1
    PLL_Div_0 = &H1F
    PLL_Div_1 = &H1
    PLL_Div_2 = &H7
    PLL_Div_3 = &H5
    Production_Line_ID = &HC2
    TSEN = 1
    TSEN_CAL_FACTOR_TRIM = 0
    TSEN_EN = 0
    TSEN_OFST = 0
    Volt_Trim = &HB
    Wafer_Size = &H1
    Wafer_Type = &H4
    WF_Good = &H1
    X_Wafer_Coordinate = &HB3
    Y_Wafer_Coordinate = &HC1
    
    
    sldata = &H8C&
    Call I2CEFuseWrite(&H1E&, sldata)
    'Call I2CEFuseRead(&H1E&, slData)
    'Debug.Print "0x1E=", slData(0) 'print 0x1E

    sldata = &H9F&
    Call I2CEFuseWrite(&H1F&, sldata)
    'Call I2CEFuseRead(&H1F&, aslReadData(1))
    'Debug.Print "0x1F=", slData(0) 'print 0x1F

    sldata = &H7E&
    Call I2CEFuseWrite(&H1D&, sldata)
    'Call I2CEFuseRead(&H1D&, slData)
    'Debug.Print "0x1D=", slData(0) 'print 0x1D
    
    'Reset all efuse as 0
    For lI = 0 To &H23
        sldata = 0
        Call I2CEFuseWrite(&H7F00& + lI, sldata)
    Next
    
    'enable efuse programming
    sldata = 0
    Call I2CEFuseWrite(&H1D&, sldata)
    
    Dim vtg_Trim_code_1 As New SiteLong
    Dim vtg_Trim_code_2 As New SiteLong
    
    'slFreq_TrimCode = &H30
    
    vtg_Trim_code_1 = slVoltage_TrimCode.BitwiseAnd(&HFF)
    vtg_Trim_code_2 = slVoltage_TrimCode.BitwiseAnd(&H3FFF).ShiftRight(8)


    I2CEFuseWrite X_Coordinate, X_Wafer_Coordinate
    
    
    

    I2CEFuseWrite Y_Coordinate, Y_Wafer_Coordinate
    I2CEFuseWrite LOT_ID_0, Device_ID1
    I2CEFuseWrite LOT_ID_1, Device_ID2
    I2CEFuseWrite LOT_ID_2, Device_ID3
    I2CEFuseWrite LOT_ID_3, Device_ID4
    I2CEFuseWrite Wafer_ID, Wafer_Type.ShiftLeft(4).Add(Wafer_Size.ShiftLeft(1)).Add(WF_Good)
    I2CEFuseWrite Manufacture_ID_1, Manufacture_week
    I2CEFuseWrite Manufacture_ID_2, Manufacture_Year
    I2CEFuseWrite Prod_SN, Production_Line_ID
    I2CEFuseWrite PLL_TRIM_0, PLL_Div_1.ShiftLeft(5).Add(PLL_Div_0)
    I2CEFuseWrite PLL_TRIM_1, PLL_Div_3.ShiftLeft(5).Add(PLL_Div_2)
    I2CEFuseWrite OSC_TRIM, slFreq_TrimCode
    I2CEFuseWrite ADC_TRIM, ADC_Gain_CMPST.ShiftLeft(2).Add(ADC_Offset_CMPST)
    I2CEFuseWrite DAC_TRIM, DAC_Gain_CMPST.ShiftLeft(2).Add(DAC_Offset_CMPST)
    I2CEFuseWrite LDO_TRIM_0, LDO_1_VSEL_Code.ShiftLeft(5).Add(LDO_1_VCMP_Code.ShiftLeft(2).Add(LDO_1_VOFST_Code))
    I2CEFuseWrite LDO_TRIM_1, LDO_2_VSEL_Code.ShiftLeft(5).Add(LDO_2_VCMP_Code.ShiftLeft(2).Add(LDO_2_VOFST_Code))
    I2CEFuseWrite LDO_TRIM_2, LDO_3_VSEL_Code.ShiftLeft(5).Add(LDO_3_VCMP_Code.ShiftLeft(2).Add(LDO_3_VOFST_Code))
    I2CEFuseWrite BUCK_TRIM_0, BUCK_1_VSEL_Code.ShiftLeft(5).Add(BUCK_1_ISET_Code)
    I2CEFuseWrite BUCK_TRIM_1, BUCK_2_VSEL_Code.ShiftLeft(5).Add(BUCK_2_ISET_Code)
    I2CEFuseWrite BUCK_TRIM_2, BUCK_3_VSEL_Code.ShiftLeft(5).Add(BUCK_3_ISET_Code)
    I2CEFuseWrite BUBO_TRIM_0, BUCK_JUMP_SIZE_0.ShiftLeft(4).Add(OSC_TRIM_0)
    I2CEFuseWrite BUBO_TRIM_1, BUCK_JUMP_SIZE_1.ShiftLeft(4).Add(OSC_TRIM_1)
    I2CEFuseWrite Vtg_TRIM_1, vtg_Trim_code_1
    I2CEFuseWrite Vtg_TRIM_2, vtg_Trim_code_2
    I2CEFuseWrite CPU_TRIM, CPU_MAX_FREQ
    
    slReadBackExpect(0) = X_Wafer_Coordinate
    slReadBackExpect(1) = Y_Wafer_Coordinate
    slReadBackExpect(2) = Device_ID1
    slReadBackExpect(3) = Device_ID2
    slReadBackExpect(4) = Device_ID3
    slReadBackExpect(5) = Device_ID4
    slReadBackExpect(6) = Wafer_Type.ShiftLeft(4).Add(Wafer_Size.ShiftLeft(1)).Add(WF_Good)
    slReadBackExpect(7) = Manufacture_week
    slReadBackExpect(8) = Manufacture_Year
    slReadBackExpect(9) = Production_Line_ID
    slReadBackExpect(10) = &HFF&
    slReadBackExpect(11) = &HFF&
    slReadBackExpect(12) = PLL_Div_1.ShiftLeft(5).Add(PLL_Div_0)
    slReadBackExpect(13) = PLL_Div_3.ShiftLeft(5).Add(PLL_Div_2)
    slReadBackExpect(14) = slFreq_TrimCode.BitwiseOr(&H80)
    slReadBackExpect(15) = ADC_Gain_CMPST.ShiftLeft(2).Add(ADC_Offset_CMPST)
    slReadBackExpect(16) = DAC_Gain_CMPST.ShiftLeft(2).Add(DAC_Offset_CMPST)
    slReadBackExpect(17) = LDO_1_VSEL_Code.ShiftLeft(5).Add(LDO_1_VCMP_Code.ShiftLeft(2).Add(LDO_1_VOFST_Code))
    slReadBackExpect(18) = LDO_2_VSEL_Code.ShiftLeft(5).Add(LDO_2_VCMP_Code.ShiftLeft(2).Add(LDO_2_VOFST_Code))
    slReadBackExpect(19) = LDO_3_VSEL_Code.ShiftLeft(5).Add(LDO_3_VCMP_Code.ShiftLeft(2).Add(LDO_3_VOFST_Code))
    slReadBackExpect(20) = BUCK_1_VSEL_Code.ShiftLeft(5).Add(BUCK_1_ISET_Code)
    slReadBackExpect(21) = BUCK_2_VSEL_Code.ShiftLeft(5).Add(BUCK_2_ISET_Code)
    slReadBackExpect(22) = BUCK_3_VSEL_Code.ShiftLeft(5).Add(BUCK_3_ISET_Code)
    slReadBackExpect(23) = BUCK_JUMP_SIZE_0.ShiftLeft(4).Add(OSC_TRIM_0)
    slReadBackExpect(24) = BUCK_JUMP_SIZE_1.ShiftLeft(4).Add(OSC_TRIM_1)
    slReadBackExpect(25) = vtg_Trim_code_1
    slReadBackExpect(26) = vtg_Trim_code_2.BitwiseOr(&H80)
    slReadBackExpect(27) = CPU_MAX_FREQ


    sldata = &HFF&
    I2CEFuseWrite Trim_Byte_0, sldata
    I2CEFuseWrite Trim_Byte_1, sldata

    thehdw.Wait 50 * ms
    
    Call I2CEFuseRead(X_Coordinate, slReadback(0), slReadBackExpect(0)) 'dummy read
    Call I2CEFuseRead(Y_Coordinate, slReadback(1), slReadBackExpect(1))
    Call I2CEFuseRead(LOT_ID_0, slReadback(2), slReadBackExpect(2))
    Call I2CEFuseRead(LOT_ID_1, slReadback(3), slReadBackExpect(3))
    Call I2CEFuseRead(LOT_ID_2, slReadback(4), slReadBackExpect(4))
    Call I2CEFuseRead(LOT_ID_3, slReadback(5), slReadBackExpect(5))
    Call I2CEFuseRead(Wafer_ID, slReadback(6), slReadBackExpect(6))
    Call I2CEFuseRead(Manufacture_ID_1, slReadback(7), slReadBackExpect(7))
    Call I2CEFuseRead(Manufacture_ID_2, slReadback(8), slReadBackExpect(8))
    Call I2CEFuseRead(Prod_SN, slReadback(9), slReadBackExpect(9))
    Call I2CEFuseRead(Trim_Byte_0, slReadback(10), slReadBackExpect(10))
    Call I2CEFuseRead(Trim_Byte_1, slReadback(11), slReadBackExpect(11))
    Call I2CEFuseRead(PLL_TRIM_0, slReadback(12), slReadBackExpect(12))
    Call I2CEFuseRead(PLL_TRIM_1, slReadback(13), slReadBackExpect(13))
    Call I2CEFuseRead(OSC_TRIM, slReadback(14), slReadBackExpect(14))
    Call I2CEFuseRead(ADC_TRIM, slReadback(15), slReadBackExpect(15))
    Call I2CEFuseRead(DAC_TRIM, slReadback(16), slReadBackExpect(16))
    Call I2CEFuseRead(LDO_TRIM_0, slReadback(17), slReadBackExpect(17))
    Call I2CEFuseRead(LDO_TRIM_1, slReadback(18), slReadBackExpect(18))
    Call I2CEFuseRead(LDO_TRIM_2, slReadback(19), slReadBackExpect(19))
    Call I2CEFuseRead(BUCK_TRIM_0, slReadback(20), slReadBackExpect(20))
    Call I2CEFuseRead(BUCK_TRIM_1, slReadback(21), slReadBackExpect(21))
    Call I2CEFuseRead(BUCK_TRIM_2, slReadback(22), slReadBackExpect(22))
    Call I2CEFuseRead(BUBO_TRIM_0, slReadback(23), slReadBackExpect(23))
    Call I2CEFuseRead(BUBO_TRIM_1, slReadback(24), slReadBackExpect(24))
    Call I2CEFuseRead(Vtg_TRIM_1, slReadback(25), slReadBackExpect(25))
    Call I2CEFuseRead(Vtg_TRIM_2, slReadback(26), slReadBackExpect(26))
    Call I2CEFuseRead(CPU_TRIM, slReadback(27), slReadBackExpect(27))

    For lI = 0 To 27
         
        TheExec.Flow.TestLimit slReadback(lI), slReadBackExpect(lI), slReadBackExpect(lI), forceresults:=tlForceFlow
        TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex - 1
    Next
    
    
    

    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function EFuseI2CPortTestDSSC()
    On Error GoTo errHandler

    Dim sldata As New SiteLong
    Dim i As Long
    Dim aslReadData(2) As New SiteLong
    
    

    ' Load level and timing.
    thehdw.Digital.ApplyLevelsTiming True, True, True, tlPowered

    sldata = &H8C&
    Call I2CEFuseWrite(&H1E&, sldata)
    Call I2CEFuseRead(&H1E&, aslReadData(0), sldata)
    'Debug.Print "0x1E=", slData(0) 'print 0x1E

    sldata = &H9F&
    Call I2CEFuseWrite(&H1F&, sldata)
    Call I2CEFuseRead(&H1F&, aslReadData(1), sldata)
    'Debug.Print "0x1F=", slData(0) 'print 0x1F

    sldata = &H7E
    Call I2CEFuseWrite(&H1D&, sldata)
    Call I2CEFuseRead(&H1D&, aslReadData(2), sldata)
    'Debug.Print "0x1D=", slData(0) 'print 0x1D
    
    TheExec.Flow.TestLimit aslReadData(0), forceresults:=tlForceFlow
    TheExec.Flow.TestLimit aslReadData(1), forceresults:=tlForceFlow
    TheExec.Flow.TestLimit aslReadData(2), forceresults:=tlForceFlow
        


    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function EFusePgmDSSC()
    On Error GoTo errHandler

    Dim sldata As New SiteLong
    Dim i As Long
    Dim aslReadData(6) As New SiteLong
    
    ' Load level and timing.
    thehdw.Digital.ApplyLevelsTiming True, True, True, tlPowered

    sldata = &H8C&
    Call I2CEFuseWrite(&H1E&, sldata)
    Call I2CEFuseRead(&H1E&, aslReadData(0), sldata)
    'Debug.Print "0x1E=", slData(0) 'print 0x1E

    sldata = &H9F&
    Call I2CEFuseWrite(&H1F&, sldata)
    Call I2CEFuseRead(&H1F&, aslReadData(1), sldata)
    'Debug.Print "0x1F=", slData(0) 'print 0x1F

    sldata = &H7E&
    Call I2CEFuseWrite(&H1D&, sldata)
    Call I2CEFuseRead(&H1D&, aslReadData(2), sldata)
    'Debug.Print "0x1D=", slData(0) 'print 0x1D
    
    sldata = 0
    Call I2CEFuseWrite(&H7F00&, sldata)
    Call I2CEFuseRead(&H7F00&, aslReadData(3), sldata)
    
    

    sldata = 0
    Call I2CEFuseWrite(&H1D&, sldata)
    Call I2CEFuseRead(&H1D&, aslReadData(4), sldata)
    'Debug.Print "0x1D=", slData(0) 'print 0x1D

    sldata = &HAA&
    Call I2CEFuseWrite(&H7F00&, sldata)
    Call I2CEFuseRead(&H7F00&, aslReadData(5), sldata)


    sldata = &H55&
    Call I2CEFuseWrite(&H7F00&, sldata, False)
    sldata = &HFF&
    Call I2CEFuseRead(&H7F00&, aslReadData(6), sldata)

    
    TheExec.Flow.TestLimit aslReadData(0), forceresults:=tlForceFlow
    TheExec.Flow.TestLimit aslReadData(1), forceresults:=tlForceFlow
    TheExec.Flow.TestLimit aslReadData(2), forceresults:=tlForceFlow
    TheExec.Flow.TestLimit aslReadData(3), forceresults:=tlForceFlow
    TheExec.Flow.TestLimit aslReadData(4), forceresults:=tlForceFlow
    TheExec.Flow.TestLimit aslReadData(5), forceresults:=tlForceFlow
    TheExec.Flow.TestLimit aslReadData(6), forceresults:=tlForceFlow
    
    
    
        


    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function EFusePgmPA()
    On Error GoTo errHandler

    Dim sldata As New SiteLong
    Dim i As Long
    Dim aslReadData(6) As New SiteLong
    
    ' Load level and timing.
    thehdw.Digital.ApplyLevelsTiming True, True, True, tlPowered

    sldata = &H8C&
    Call I2CEFuseWritePA(164, &H1E&, sldata)
    Call I2CEFuseReadPA(164, &H1E&, 165, aslReadData(0))
    'Debug.Print "0x1E=", slData(0) 'print 0x1E

    sldata = &H9F&
    Call I2CEFuseWritePA(164, &H1F&, sldata)
    Call I2CEFuseReadPA(164, &H1F&, 165, aslReadData(1))
    'Debug.Print "0x1F=", slData(0) 'print 0x1F

    sldata = &H7E&
    Call I2CEFuseWritePA(164, &H1D&, sldata)
    Call I2CEFuseReadPA(164, &H1D&, 165, aslReadData(2))
    'Debug.Print "0x1D=", slData(0) 'print 0x1D
    
    sldata = 0
    Call I2CEFuseWritePA(164, &H7F00&, sldata)
    'thehdw.Wait 200 * ms
    Call I2CEFuseReadPA(164, &H7F00&, 165, aslReadData(3))
    Call I2CEFuseReadPA(164, &H7F00&, 165, aslReadData(3)) ' the second read is necessary since FPGA need time to write&read EEPROM! read twice can garantee reading is successful
    
    sldata = 0
    Call I2CEFuseWritePA(164, &H1D&, sldata)
    Call I2CEFuseReadPA(164, &H1D&, 165, aslReadData(4))
    'Debug.Print "0x1D=", slData(0) 'print 0x1D
    
    sldata = &HAA&
    Call I2CEFuseWritePA(164, &H7F00&, sldata)
    'thehdw.Wait 200 * ms
    Call I2CEFuseReadPA(164, &H7F00&, 165, aslReadData(5))
    Call I2CEFuseReadPA(164, &H7F00&, 165, aslReadData(5)) ' the second read is necessary since FPGA need time to write&read EEPROM! read twice can garantee reading is successful
    
    sldata = &H55&
    Call I2CEFuseWritePA(164, &H7F00&, sldata)
    'thehdw.Wait 200 * ms
    Call I2CEFuseReadPA(164, &H7F00&, 165, aslReadData(6))
    Call I2CEFuseReadPA(164, &H7F00&, 165, aslReadData(6)) ' the second read is necessary since FPGA need time to write&read EEPROM! read twice can garantee reading is successful
    
    TheExec.Flow.TestLimit aslReadData(0), forceresults:=tlForceFlow
    TheExec.Flow.TestLimit aslReadData(1), forceresults:=tlForceFlow
    TheExec.Flow.TestLimit aslReadData(2), forceresults:=tlForceFlow
    TheExec.Flow.TestLimit aslReadData(3), forceresults:=tlForceFlow
    TheExec.Flow.TestLimit aslReadData(4), forceresults:=tlForceFlow
    TheExec.Flow.TestLimit aslReadData(5), forceresults:=tlForceFlow
    TheExec.Flow.TestLimit aslReadData(6), forceresults:=tlForceFlow
    
    
    
        


    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function EFuseI2CPortTestPA()
    On Error GoTo errHandler

    Dim sldata As New SiteLong
    Dim i As Long
    Dim aslReadData(2) As New SiteLong
    
    

    ' Load level and timing.
    thehdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    
    thehdw.Protocol.Ports("I2C_EFUSE_PA").Enabled = True
    
    sldata = &H8C&
    Call I2CEFuseWritePA(164, &H1E&, sldata)
    Call I2CEFuseReadPA(164, &H1E&, 164 + 1, aslReadData(0))
    'Debug.Print "0x1E=", slData(0) 'print 0x1E

    sldata = &H9F&
    Call I2CEFuseWritePA(164, &H1F&, sldata)
    Call I2CEFuseReadPA(164, &H1F&, 164 + 1, aslReadData(1))
    'Debug.Print "0x1F=", slData(0) 'print 0x1F

    sldata = &H7E
    Call I2CEFuseWritePA(164, &H1D&, sldata)
    Call I2CEFuseReadPA(164, &H1D&, 164 + 1, aslReadData(2))
    'Debug.Print "0x1D=", slData(0) 'print 0x1D
    
    TheExec.Flow.TestLimit aslReadData(0), forceresults:=tlForceFlow
    TheExec.Flow.TestLimit aslReadData(1), forceresults:=tlForceFlow
    TheExec.Flow.TestLimit aslReadData(2), forceresults:=tlForceFlow


    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function
Private Function EFuseI2CPAClearAllModule()
    Dim iIndex As Long
    If thehdw.Protocol.Ports("I2C_EFUSE_PA").Modules.IsRecorded("EFuseClearAll", False, False, True, True) = False Then
        For iIndex = &H7F00& To &H7FFF&
            With thehdw.Protocol.Ports("I2C_EFUSE_PA")
                .NWire.Frames("I2C_Write_Efuse").Fields("W_DeviceID").Value = 164
                .NWire.Frames("I2C_Write_Efuse").Fields("W_RegAddr").Value = iIndex
                .NWire.Frames("I2C_Write_Efuse").Fields("W_Data").Value = 170
                .NWire.Frames("I2C_Write_Efuse").Execute
                '.IdleWait
            End With
        Next iIndex
        thehdw.Protocol.Ports("I2C_EFUSE_PA").Modules.StopRecording
    End If
End Function
'**************************************************************************************************'
' Procedure    : I2CWrite
' Author       : Tim Guo
' DateTime     : Oct.2019
' Purpose      : Write an I2C Efuse register
' Example      : I2CWrite Addr, slValues
' Input        : Register address, register contents (multisite)
' Return Value :
'**************************************************************************************************'
Public Function I2CEFuseWrite(Addr As eI2CEFuseRegs, slvalues As SiteLong, Optional bValidateWrite As Boolean = True)
    On Error GoTo errHandler

    Dim dData(2) As Double
    Dim lData As Long
    Dim vSite As Variant
    Dim dspwData As New DSPWave
    Dim PatternName As String
    Dim slReadback As New SiteLong
    Dim lI As Long

    PatternName = ".\Patterns\I2C_EFUSE_Write_slow.pat"

    thehdw.Patterns(PatternName).Load
    'Call TheExec.Datalog.WriteComment(PatternName)
    thehdw.DSSC.Pins("I2C_SDA_EFUSE").Pattern(PatternName).Source.Signals.Add ("I2C_SendData")
    thehdw.DSSC.Pins("I2C_SDA_EFUSE").Pattern(PatternName).Source.Signals.DefaultSignal = "I2C_SendData"

    For Each vSite In TheExec.Sites
        lData = 0
        dData(0) = (Addr And &HFF00&) / &H100&
        dData(1) = Addr And &HFF&
        dData(2) = slvalues
        'Call TheExec.Datalog.WriteComment(dData(0) & " " & dData(1) & " " & dData(2))
        'Creation of the Segment for the Dig_src
        Call TheExec.WaveDefinitions.CreateWaveDefinition("I2C_SendData" & vSite, dData, True)

        With thehdw.DSSC.Pins("I2C_SDA_EFUSE").Pattern(PatternName).Source.Signals
            .Item("I2C_SendData").WaveDefinitionName = "I2C_SendData" & vSite
            .Item("I2C_SendData").LoadSamples
        End With
    Next vSite
    
    For lI = 1 To 10
        thehdw.Patterns(PatternName).Start
        'Wait until the pattern halts
        Call thehdw.Digital.Patgen.HaltWait
        
        If bValidateWrite And TheExec.TesterMode = testModeOnline Then
            Call I2CEFuseRead(Addr, slReadback, slvalues, False)
            If slReadback.BitwiseXor(slvalues).Compare(EqualTo, 0).All(True) = True Then 'And TheExec.TesterMode = testModeOnline Then
                lI = 10000000
            Else
                Debug.Print Addr, slReadback(0), slvalues(0)
            End If
        End If
    Next
    
    Exit Function

errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function


'**************************************************************************************************'
' Procedure    : I2CEfuseRead
' Author       : Tim Guo
' DateTime     : Oct.2019
' Purpose      : Read an SPI register
' Example      : I2CRead eRD_DEVICE_ID, slResults
' Input        : Register address
' Return Value : Register contents in siteLong (multisite)
'**************************************************************************************************'
Public Sub I2CEFuseRead(ByVal Addr As eI2CEFuseRegs, ByRef slResults As SiteLong, ByVal matchValue As SiteLong, Optional bPringLog As Boolean = True)
    On Error GoTo errHandler

    Dim dData(1) As Double
    Dim dspResults As New DSPWave
    'Dim slResults As New SiteLong
    Dim vSite As Variant
    Dim PatternName As String
    Dim lI As Long
    Dim slLastRead As New SiteLong
    Dim sbEnable As New SiteBoolean

   
    PatternName = ".\Patterns\I2C_EFUSE_Read_slow.pat"

    dData(0) = (Addr And &HFF00&) / &H100&
    dData(1) = Addr And &HFF&
    
    'Creation of the Segment for the Dig_src
    Call TheExec.WaveDefinitions.CreateWaveDefinition("I2C_SendData", dData, True)
    
    slLastRead = -1
    
    sbEnable = True
    thehdw.Patterns(PatternName).Load
    'Call TheExec.Datalog.WriteComment(PatternName)
    For lI = 1 To 100
        With thehdw.DSSC.Pins("I2C_SDA_EFUSE").Pattern(PatternName).Source.Signals
            .Add ("I2C_SendData")
            .Item("I2C_SendData").WaveDefinitionName = "I2C_SendData"
            '.Item("I2C_SendData").SampleSize = 1
            .Item("I2C_SendData").LoadSamples
            .DefaultSignal = "I2C_SendData"
        End With
    
        'Creation of the Segment for the Dig_cap
        With thehdw.DSSC.Pins("I2C_SDA_EFUSE").Pattern(PatternName).Capture.Signals
            .Add ("SpiRcv")
            .Item("SpiRcv").SampleSize = 1
            .Item("SpiRcv").LoadSettings
        End With
    
        thehdw.Patterns(PatternName).Start
        'Wait until the pattern halts
        Call thehdw.Digital.Patgen.HaltWait
    
        'Get the results into a dsp wave
        dspResults = thehdw.DSSC.Pins("I2C_SDA_EFUSE").Pattern(PatternName).Capture.Signals.Item("SpiRcv").DSPWave
    
        For Each vSite In TheExec.Sites.Active
            'If thehdw.Digital.Patgen.PatternBurstPassed(vSite) = False Then
                'TheExec.Datalog.WriteComment (Space(20) & "I2C Write @ Addr = &h" & CStr(Hex(Addr)) & " for site" & vSite & " failed!")
            'Else
            slResults(vSite) = dspResults(vSite).data(0)
            'TheExec.Datalog.WriteComment "***Site" & vSite & ", ADDR " & Addr & " Efuse reading " & slResults(vSite) & " Expecting " & matchValue & ", " & lI
            If matchValue = slResults(vSite) And matchValue <> -1 Then
                sbEnable = False
            End If
            
            If lI >= 100 And sbEnable = True And bPringLog Then
                TheExec.Datalog.WriteComment "Site" & vSite & " Efuse reading failed!"
            ElseIf sbEnable = True And bPringLog Then
                TheExec.Datalog.WriteComment "Site" & vSite & ", ADDR " & Addr & " Efuse reading " & slResults(vSite) & " Expecting " & matchValue & ", " & lI
            End If
        Next vSite
        TheExec.Sites.Selected = sbEnable
        For Each vSite In TheExec.Sites
            If sbEnable.All(False) = True Or matchValue(vSite) = -1 Then lI = 100000
        Next vSite
        'sbEnable.
    Next
    
    TheExec.Sites.Selected = True
    
    Exit Sub

errHandler:
    If AbortTest Then Exit Sub Else Resume Next
End Sub


Public Function I2CEFuseReadPA(lWDeviceID As Long, Reg_Addr As Long, lRCmd As Long, slReadData As SiteLong) As SiteLong

    Dim ModuleName As String
    Dim Read_pld As New PinListData
    Dim ReadWave As New DSPWave
    Dim vSite As Variant

    With thehdw.Protocol.Ports("I2C_EFUSE_PA")
        .NWire.CMEM.MoveMode = tlNWireCMEMMoveMode_Databus
        Read_pld = .NWire.CMEM.DSPWave
        .NWire.Frames("I2C_Read_Efuse").Fields("W_DeviceID").Value = lWDeviceID
        .NWire.Frames("I2C_Read_Efuse").Fields("W_RegAddr").Value = Reg_Addr
        .NWire.Frames("I2C_Read_Efuse").Fields("R_CMD").Value = lRCmd
        .NWire.Frames("I2C_Read_Efuse").Execute (tlNWireExecutionType_CaptureInCMEM)
        .IdleWait

        For Each vSite In TheExec.Sites.Active
            ReadWave = Read_pld.Pins(0).Value
            slReadData(vSite) = ReadWave.ElementLite(0)
        Next vSite
    End With

End Function
Public Function I2CEFuseWritePA(lWDeviceID As Long, Reg_Addr As Long, Reg_Data As SiteLong)

    Dim ModuleName As String
    Dim Read_pld As New PinListData
    Dim ReadWave As New DSPWave
    Dim vSite As Variant

    With thehdw.Protocol.Ports("I2C_EFUSE_PA")
        .NWire.Frames("I2C_Write_Efuse").Fields("W_DeviceID").Value = lWDeviceID
        .NWire.Frames("I2C_Write_Efuse").Fields("W_RegAddr").Value = Reg_Addr
        .NWire.Frames("I2C_Write_Efuse").Fields("W_Data").Value = Reg_Data
        .NWire.Frames("I2C_Write_Efuse").Execute
        .IdleWait
    End With

End Function


