#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
DetectHiddenWindows, on

class VerbosityLevelConfig {
    value := 0
    text := ""
    
    __New(value, text) {
        this.value := value
        this.text := text
    }
}

global VERBOSITY_LEVEL := {}
VERBOSITY_LEVEL.None     := 0
VERBOSITY_LEVEL.Info     := 1
VERBOSITY_LEVEL.Warning  := 2
VERBOSITY_LEVEL.Error    := 4
VERBOSITY_LEVEL.Critical := 8
VERBOSITY_LEVEL.Debug    := 16
VERBOSITY_LEVEL.All      := 128

global VERBOSITY_LEVEL_CONFIG := {}
VERBOSITY_LEVEL_CONFIG[VERBOSITY_LEVEL.None] := new VerbosityLevelConfig(VERBOSITY_LEVEL.None, "")
VERBOSITY_LEVEL_CONFIG[VERBOSITY_LEVEL.Info] := new VerbosityLevelConfig(VERBOSITY_LEVEL.Info, "INFO")
VERBOSITY_LEVEL_CONFIG[VERBOSITY_LEVEL.Warning] := new VerbosityLevelConfig(VERBOSITY_LEVEL.Warning, "WARNING")
VERBOSITY_LEVEL_CONFIG[VERBOSITY_LEVEL.Error] := new VerbosityLevelConfig(VERBOSITY_LEVEL.Error, "ERROR")
VERBOSITY_LEVEL_CONFIG[VERBOSITY_LEVEL.Critical] := new VerbosityLevelConfig(VERBOSITY_LEVEL.Critical, "CRITICAL")
VERBOSITY_LEVEL_CONFIG[VERBOSITY_LEVEL.Debug] := new VerbosityLevelConfig(VERBOSITY_LEVEL.Debug, "DEBUG")
VERBOSITY_LEVEL_CONFIG[VERBOSITY_LEVEL.All] := new VerbosityLevelConfig(VERBOSITY_LEVEL.All, "")

global MessageBoxVerbosity := VERBOSITY_LEVEL.Debug
MessageBoxVerbosity := MessageBoxVerbosity | VERBOSITY_LEVEL.Critical

Log(level, message) {    
    ;todo log to file
    if (MessageBoxVerbosity >= VERBOSITY_LEVEL.All || MessageBoxVerbosity & level > 0) {
        MsgBox % VERBOSITY_LEVEL_CONFIG[level].text . ": " . message
    }
}

LogInfo(message) {
    Log(VERBOSITY_LEVEL.Info, message)
}

LogWarning(message) {
    Log(VERBOSITY_LEVEL.Warning, message)
}

LogError(message) {
    Log(VERBOSITY_LEVEL.Error, message)
}

LogCritical(message) {
    Log(VERBOSITY_LEVEL.Critical, message)
}

LogDebug(message) {
    Log(VERBOSITY_LEVEL.Debug, message)
}

;Switches virtual line output between 2 different external devices
;Uses KS (kernal streaming) audio repeater
;   Harder to configure
;   Most of the time only 1 external output device is allowed (limitation can be bypassed by have a static Virtual Input redirect to the virtual output line).
;   Lowest latency.

;path to VAC audio repepater
global appPath := "C:\Program Files\Virtual Audio Cable\audiorepeater_ks.exe"

;the iconNum is only used
global VAC_ICONS := {}
VAC_ICONS.default := {file: "mmres.dll", iconNum: 4}
VAC_ICONS.speakers := {file: "speaker-white.ico"}
VAC_ICONS.headphones := {file: "headset-white.ico"}
VAC_ICONS.bluetooth := {file: "bluetooth-white.ico"}

;probably won't have to change these
global defaultSettings := {}
defaultSettings.samplingRate := 48000
defaultSettings.bitsPerSample := 16
defaultSettings.numChannels := 2
defaultSettings.channelConfig := "Stereo"
defaultSettings.totalBuffer := 20
defaultSettings.numBuffers := 8
defaultSettings.cpuPriority := "High" ; Normal, High, Realtime
defaultSettings.autoStart := false
defaultSettings.channelConfigIn := "Stereo"
defaultSettings.channelConfigOut := "Stereo"
defaultSettings.continuousQueueIn := false
defaultSettings.continuousQueueOut := false
defaultSettings.outputPrefillPercent := 50
defaultSettings.winTarget := "Min" ;Set to "" (normal), "Min" (minimizes to tray at startup), "Hide" (Hidden from taskbar and tray)
defaultSettings.icon := VAC_ICONS.default

class VACParam {
    paramName := ""
    paramType := "string"
    
    __New(paramName, paramType) {
        this.paramName := paramName
        this.paramType := paramType
    }
    
    BuildParam(value) {
        param := ""
        if (value != "" || this.paramType == "string") {
            if (this.paramType != "boolean" || value == true) {
                param := "/" . this.paramName                
                if (this.paramType != "boolean") {
                    if (this.paramType == "string") {
                        value := """" . value . """"
                    }
                    param := param . ":" . value
                }
            }
        }
        return param
    }    
}

;defines the command line param config (DON'T CHANGE)
global VAC := {}
VAC.winName              := new VACParam("WindowName"   , "string" )
VAC.input                := new VACParam("Input"        , "string" )
VAC.output               := new VACParam("Output"       , "string" )
VAC.samplingRate         := new VACParam("SamplingRate" , "number" )
VAC.bitsPerSample        := new VACParam("BitsPerSample", "number" )
VAC.numChannels          := new VACParam("Channels"     , "number" )
VAC.channelConfig        := new VACParam("ChanCfg"      , "string" ) 
VAC.totalBuffer          := new VACParam("BufferMs"     , "number" )  
VAC.numBuffers           := new VACParam("Buffers"      , "number" )
VAC.cpuPriority          := new VACParam("Priority"     , "string" )
VAC.channelConfigIn      := new VACParam("ChanCfgIn"    , "string" )
VAC.channelConfigOut     := new VACParam("ChanCfgOut"   , "string" )
VAC.continuousQueueIn    := new VACParam("ContQueueIn"  , "boolean")
VAC.continuousQueueOut   := new VACParam("ContQueueOut" , "boolean")
VAC.outputPrefillPercent := new VACParam("OutputPreFill", "number" )
VAC.autoStart            := new VACParam("AutoStart"    , "boolean")
VAC.closeInstance        := new VACParam("CloseInstance", "string" )

;convienence function to change the tray icon
SetTrayIcon(options) {
    if (options.HasKey("file")) {
        if (options.HasKey("iconNum")) {        
            Menu, Tray, Icon, % options.file, % options.iconNum
        } else {
            Menu, Tray, Icon, % options.file
        }
    }
}

;sets the icon the default one
SetTrayIcon(VAC_ICONS.default)

;class to store individual settings for each repeater started
;you can override default settings
class VACRepeaterSettings {
    output := ""
    
    ;constructor
    __New(output) {
        this.output := output
    }
    
    __Get(key) {
        return defaultSettings[key]
    }
}

class ParamBuilder {
    settings := {}
    params := ""
    
    __New(settings) {
        this.settings := settings
    }
    
    Add(paramName) {
        paramString := VAC[paramName].BuildParam(this.settings[paramName])
        ;LogDebug("Adding=" . paramName . " :: ParamString=" . paramString)
        if (paramString != "") {
            if (this.params != "") {
                this.params := this.params . " "
            }
            this.params := this.params . paramString
        }
    }
}    

class VACRepeater {
    winName := ""
    input := ""
    settings := {}
    uid := -1
    pid := -1
    updateTrayIcon := false
    
    startParams := ""
    stopParams := ""
    
    __New(winName, input, settings, updateTrayIcon) {
        this.winName := winName
        this.closeInstance := winName
        this.input := input
        this.settings := settings
        this.updateTrayIcon := updateTrayIcon
    }
    
    __Get(setting) {
        return this.settings[setting]
    }
    
    Open() {
        if (this.startParams == "") {
            pb := new ParamBuilder(this)
            pb.Add("winName")
            pb.Add("input")
            pb.Add("output")
            pb.Add("samplingRate")
            pb.Add("bitsPerSample")
            pb.Add("numChannels")
            pb.Add("channelConfig")
            pb.Add("totalBuffer")
            pb.Add("numBuffers")
            pb.Add("cpuPriority")
            pb.Add("channelConfigIn")
            pb.Add("channelConfigOut")
            pb.Add("continuousQueueIn")
            pb.Add("continuousQueueOut")
            pb.Add("outputPrefillPercent")
            pb.Add("autoStart")
            this.startParams := pb.params
            pb := {}            
        }        
        if (!WinExist(this.winName)) {
            LogInfo("Starting repeater with params: " . this.startParams)            
            target := """" . appPath . """ " . this.startParams
            Run, %target%, , % this.settings.winTarget, pid
            this.pid := pid
            WinWait, % this.winName, , 3
            this.Pulse()
        }
    }
        
    Close() {
        if (this.stopParams == "") {
            pb := new ParamBuilder(this)
            pb.Add("closeInstance")
            this.stopParams := pb.params
            pb := {}            
        } 
        if (WinExist(this.winName)) {
            LogInfo("Stopping repeater with params: " . this.stopParams)
            target := """" . appPath . """ " . this.stopParams
            Run, %target%, , Hide
            WinWaitClose, , , 3
        }
    }
    
    Start() {
        if (!WinExist(this.winName)) {
            this.Open()
        }
        ControlClick, Start, % this.winName
        if (this.updateTrayIcon) {
            SetTrayIcon(this.settings.icon)
        }
    }
    
    Stop() {
        if (WinExist(this.winName)) {
            ControlClick, Stop, % this.winName
        }        
        SetTrayIcon(VAC_ICONS.default)
    }
    
    Restart() {
        this.Stop()
        this.Start()
    }
    
    Activate() {
        if (WinExist(this.winName)) {
            WinActivate
            WinWaitActive, , , 3
        }
    }        
    
    Hide() {
        if (WinExist(this.winName)) {
            WinHide
        }
    }
    
    Pulse() {        
        if (WinExist(this.winName)) {
            WinActivate
            ;WinWaitActive, , , 3
            WinHide
            ;WinWaitNotActive, , , 3
        }
    }
    
    Focus() {
        if (WinExist(this.winName)) {
            ControlFocus, Start
            ControlFocus, Stop
        }
    }
            
        
}   

class WrapAroundIndex {
    index := 0
    lowerBound := 0
    upperBound := 0
    
    SetIndex(index) {
        if (index > this.upperBound) {
            index := this.lowerBound
        } else if (index < this.lowerBound) {
            index := this.upperBound
        }
        this.index := index
    }
    
    Increment(inc) {
        this.SetIndex(this.index + inc)
    }
    
    __New(lowerBound, upperBound, initialIndex) {
        if (lowerBound <= upperBound) {
            this.lowerBound := lowerBound
            this.upperBound := upperBound
        }
        this.SetIndex(initialIndex)
    }
}

;Class that defines the switchers used 
;you can define a 1->1 repeater, and also setting it to persistent
class VACSwitcher {    
    name := ""
    inputs := {}
    outputs := {} ;first one in the list is the first one initialized at startup
    persistent := false ;if true, then the shortcut will not switch to the next ouput.  If only one output and set to false, it effectively restarts the repeater everytime the shortcut is pressed.
    updateTrayIcon := false ;if set to true, then when the output is switched, the tray icon is updated.
    
    ;internal vars
    repeaterSets := {}
    setIndex := {}
    paused := false
    
    __New(name) {
        this.name := name
    }
    
    Init() {        
        if (this.inputs.Length() > 0 && this.outputs.Length() > 0) {         
            Loop % this.outputs.Length() {
                outputIndex := A_Index
                set := {}
                Loop % this.inputs.Length() {
                    inputIndex := A_Index
                    winName := "VAC (" . this.name . "): Input " . inputIndex . " - Output " . outputIndex
                    repeater := new VACRepeater(winName, this.inputs[inputIndex], this.outputs[outputIndex], this.updateTrayIcon)
                    set.Push(repeater)
                }
                this.repeaterSets.Push(set)
            }
        }
        this.setIndex := new WrapAroundIndex(1, this.repeaterSets.Length(), 1)
    }
    
    MaintainAllSets(op) {
        Loop % this.repeaterSets.Length() {
            set := this.repeaterSets[A_Index]
            Loop % set.Length() {
                repeater := set[A_Index]
                if (op == "open") {
                    repeater.Open()
                } else if (op == "close") {
                    repeater.Close()
                } else if (op == "activate") {
                    repeater.activate()
                }
            }
        }
    }
    MaintainCurrentSet(op) {
        set := this.repeaterSets[this.setIndex.index]
        Loop % set.Length() {
            repeater := set[A_Index]
            if (op == "start") {
                repeater.Start()
            } else if (op == "stop") {
                repeater.Stop()
            } else if (op == "restart") {
                repeater.Restart()
            } 
        }
    }
    
    Open() {
        this.MaintainAllSets("open")
    }
    
    Close() {
        this.MaintainAllSets("close")
    }
    
    Start() {
        this.MaintainCurrentSet("start")
    }
    
    Stop() {
        this.MaintainCurrentSet("stop")
    }
    
    Restart() {
        this.Stop()
        this.Start()
    }
    
    Pause() { 
        this.paused := true
        this.Stop()
    }
    
    Resume() {
        this.Start()
        this.paused := false
    }
    
    Switch(inc) {
        if (this.paused) {
            this.Resume()
        } else if (!this.persistent) {
            this.Stop()
            this.setIndex.Increment(inc)
            this.Start()
        } else {            
            this.Start()
        }
    }
    
    Activate() {
        this.MaintainAllSets("activate")
    }
}

;NOTE:  For KS, you need use the names in the Audio Repeater KS application, not the control panel sound devices.
;virtual lines created (note: you will get an error if the output device can only support one input).

;stores the repeater setups
global switchers := {}

;Default Device -> Switching Repeater
;For me i need this so i can use dxtory to seperate out audio streams
switcher := new VACSwitcher("Default Device")
switcher.persistent := true
switcher.inputs.Push("Virtual Cable 1")
output := new VACRepeaterSettings("Virtual Cable 3")
output.totalBuffer := 70
switcher.outputs.Push(output)
switchers.Push(switcher)

;Default comm device -> Switching repeater
;need this to capture VOIP stream seperately
switcher := new VACSwitcher("Default Comm")
switcher.persistent := true
switcher.inputs.Push("Virtual Cable 2")
output := new VACRepeaterSettings("Virtual Cable 3")
output.totalBuffer := 70
switcher.outputs.Push(output)
switchers.Push(switcher)

;;;START MIXER
;outputs the streams to the correct external output device
switcher := new VACSwitcher("Mixer")
switcher.updateTrayIcon := true
switcher.inputs.Push("Virtual Cable 3")

;speakers
output := new VACRepeaterSettings("Realtek HD Audio output")
output.totalBuffer := 20
output.icon := VAC_ICONS.speakers
switcher.outputs.Push(output)

;headphones
;output := new VACRepeaterSettings("USB Audio Device") ;usb
output := new VACRepeaterSettings("Realtek HDA SPDIF Optical Out") ;optical
output.icon := VAC_ICONS.headphones
switcher.outputs.Push(output)
switchers.Push(switcher)

;Bluetooth speakers
output := new VACRepeaterSettings("Unknown")
output.totalBuffer := 20
output.samplingRate := 44100
output.icon := VAC_ICONS.bluetooth
switcher.outputs.Push(output)
;;;END Mixer

if (switchers.Length() == 0) {
    LogCritical("No switchers defined!")
    ExitApp 1
}

MaintainSwitchers(op, inc) {
    Loop % switchers.Length() {
        switcher := switchers[A_Index]
        if (op == "start") {
            switcher.Start()
        } else if (op == "stop") {
            switcher.Stop()
        } else if (op == "restart") {
            switcher.Restart()
        } else if (op == "switch") {
            switcher.Switch(inc)
        } else if (op == "pause") {
            switcher.Pause()
        } else if (op == "resume") {
            switcher.Resume()
        } else if (op == "init") {
            switcher.Init()
        } else if (op == "open") {
            switcher.Open()
        } else if (op == "close") {
            switcher.Close()
        } else if (op == "activate") {
            switcher.Activate()
        }
    }
}

MaintainSwitchers("init", 0)
MaintainSwitchers("open", 0)
;MaintainSwitchers("activate", 0)
MaintainSwitchers("start", 0)

;ctrl+shift+F12
^+F12::
    MaintainSwitchers("pause", 0)
return 

;ctrl+shift+alt+F12
^+!F12::
    MaintainSwitchers("close", 0)
    ExitApp 0
return

;ctrl+F12
^F12::
    MaintainSwitchers("switch", 1) 
return