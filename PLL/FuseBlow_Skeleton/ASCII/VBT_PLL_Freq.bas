Attribute VB_Name = "VBT_PLL_Freq"
Dim g_TrimCode As New SiteLong

Public Function Write_TrimCode_to_Efuse_vbt() As Long
    On Error GoTo errHandler
    
    Dim sl_target_trimcode As New SiteLong
    Dim ret_trimcode As New SiteLong
    Dim Site As Variant
    
    'If this is offline, then set the global trim code to 50
    If TheExec.TesterMode = testModeOffline Then
        'g_TrimCode = 50
        For Each Site In TheExec.Sites
            g_TrimCode(Site) = 50
        Next Site
    End If
    
    'sl_target_trimcode = g_TrimCode Or &H80
    
    For Each Site In TheExec.Sites
        sl_target_trimcode(Site) = g_TrimCode(Site) Or &H80
        Call TheExec.Datalog.WriteComment("Final Trim Code for Site" & Site & " : " & g_TrimCode(Site))
    Next Site
           
    thehdw.Digital.ApplyLevelsTiming True, True, True, tlUnpowered
    thehdw.Wait 0.05
       
    'Init Efuse and  Erase Efuse data
    Call EFuseOTPErase
    
    'Write trim code
    Call I2CEFuseWrite(OSC_TRIM, sl_target_trimcode)
    
    'read from efuse
    Call I2CEFuseRead(OSC_TRIM, ret_trimcode, sl_target_trimcode)
    
    'Power Down the Device
    With thehdw.DCVI.Pins("VCCIO")
        .Voltage = 0
        .Disconnect
    End With

    With thehdw.DCVI.Pins("VCC")
        .Voltage = 0
        .Disconnect
    End With
    
    thehdw.Wait 0.5
    
'    ' =========================================================================
'    ' The below is going to be program by the students
'    ' This is to check if the Fuse blow results in the Frequency being changed
'    ' =========================================================================
'
'    ' Apply levels and timings - Power up the device
'    thehdw.Digital.ApplyLevelsTiming True, True, True, tlUnpowered
'
'    ' Enable nWire PA Engines
'    thehdw.Protocol.Ports("SPI_PORT").Enabled = True
'
'    ' Setting Up nWire HRAM£¬HRAM Capture Setup
'    thehdw.Protocol.Ports("SPI_PORT").NWire.HRAM.Setup.TriggerType = tlNWireHRAMTriggerType_Never
'    thehdw.Protocol.Ports("SPI_PORT").NWire.HRAM.Setup.WaitForEvent = False
'
'    ' Write data into register
'    SPI_write_regdata &HB, &H2
'
'    ' Set the output frequency
'    SPI_write_regdata &H2, &H1A
'
'    ' Pattern gen for MCG Clock
'    ' Load and start pattern.
'    thehdw.Patterns(PatName).Load
'    thehdw.Patterns(PatName).Start
'
'    ' Setup the Frequency Counter
'    ' Clear and reset the frequency counter.
'    Call thehdw.Digital.Pins(PintoMeasure).FreqCtr.Clear
'
'    ' Set up the frequency counter based on passed-in parameter values.
'    With thehdw.Digital.Pins(PintoMeasure).FreqCtr
'       .EventSource = evntSrc ' VOH or VOL
'       .EventSlope = evntSlope ' Positive or Negative
'       .Enable = IntervalEnable
'       .Interval = TimeInterval ' Set Period Counter Interval in seconds
'    End With
'
'    ' Start the frequency counter and read measurements
'    thehdw.Digital.Pins(PintoMeasure).FreqCtr.Start
'
'    thehdw.Wait 0.1
'
'    ' Return calculated frequency values for all sites based on the read
'    MeasFreq = thehdw.Digital.Pins(PintoMeasure).FreqCtr.MeasureFrequency
'
'    SPI_read_regdata &H5, trimcode_rdback
'
'    TheExec.Flow.TestLimit resultVal:=trimcode_rdback, forceresults:=tlForceFlow, Tname:="Trimcode_Readback"
'
'    ' Halt the pattern.
'    thehdw.Digital.Patgen.Halt
'
    
    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function
