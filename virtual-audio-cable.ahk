;Switches virtual line output between 2 different external devices
;Uses MME audio repeater
;   Easier to configure
;   Allows same output device
;   Higher latency

#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
DetectHiddenWindows, on

Menu, Tray, Icon, mmres.dll, 4

;path to VAC audio repepater
global appPath := "C:\Program Files\Virtual Audio Cable\audiorepeater.exe"

;NOTE:  For MME, you can either use the names in the Audio Repeater application or the full, non-truncated names in the control panel, sound devices.
;virtual lines created
global inputs := []
inputs.Push("Virtual Input 1 (Virtual Audio Cable)")
inputs.Push("Virtual Input 2 (Virtual Audio Cable)")

global outputs := []
outputs.Push({name: "Speakers (Realtek High Definition Audio)", icon: "speaker-white.ico"})
outputs.Push({name: "Headphones (3- SteelSeries H Wireless)", icon: "headset-white.ico"})

if (inputs.Length() == 0 || outputs.Length() == 0) {
    Msgbox % "config error"
    ExitApp 1
}

global windowNames := []
Loop % inputs.Length() {
    windowNames[A_Index] := ""
}

global outputIndex := 1

StartVAC(winName, input, output 
    , samplingRate, bitsPerSample
    , numChannels, channelConfig
    , buffer, numBuffers
    , priority, autoStart) {
    
    ;set the params
    params := ""
    params := params . " /WindowName:" . """" . winName . """"
    params := params . " /Input:" . """" . SubStr(input, 1, 31) . """"
    params := params . " /Output:" . """" . SubStr(output, 1, 31) . """"
    params := params . " /SamplingRate:" . samplingRate
    params := params . " /BitsPerSample:" . bitsPerSample
    params := params . " /Channels:" . numChannels
    params := params . " /ChanCfg:" . """" . channelConfig . """"
    params := params . " /BufferMs:" . buffer
    params := params . " /Buffers:" . numBuffers
    params := params . " /Priority:" . priority
    if (autoStart) {
        params := params . " /AutoStart"
    }
    
    Run "%appPath%" %params%, , Min, newProcessId
    return newProcessId
}

StartDefaultVAC(winName, input, output) {
    samplingRate := 44100
    bitsPerSample := 16
    numChannels := 2
    channelConfig := "Stereo"
    totalBuffer := 200
    numBuffers := 16
    cpuPriority := "High"
    autoStart := true
    
    return StartVAC(winName, input, output, samplingRate, bitsPerSample, numChannels, channelConfig, totalBuffer, numBuffers, cpuPriority, autoStart)
}

CloseVAC(winName) {
    IfWinExist, %winName%
        WinClose
}

CloseAllVAC() {
    Loop % windowNames.Length() {
        winName := windowNames[A_Index]
        CloseVAC(winName)
    }
}

SwitchOutput(outputIndex) {
    if (outputIndex > 0 && outputIndex <= outputs.Length()) {
        output := outputs[outputIndex]
        Loop % inputs.Length() {
            inputIndex := A_Index        
            winName := "VAC: Input " . inputIndex . " - Output " . outputIndex
            input := inputs[inputIndex]
            StartDefaultVAC(winName, input, output["name"])
            windowNames[inputIndex] := winName
        }    
        trayIcon := output["icon"]
        Menu, Tray, Icon, %trayIcon%
    }
}

IncrementOutputIndex(inc) {
    outputIndex += inc
    if (outputIndex <= 0) {
        outputIndex := outputs.Length()
    } else if (outputIndex > outputs.Length()) {
        outputIndex := 1
    }
}

;reset everything
Loop {
    IfWinExist VAC: 
        WinClose
    else
        break
}

SwitchOutput(outputIndex)
IncrementOutputIndex(1)

^+F12::
    CloseAllVAC()
    IncrementOutputIndex(-1)
return
    
^F12::
    CloseAllVAC()
    SwitchOutput(outputIndex)
    IncrementOutputIndex(1)
return