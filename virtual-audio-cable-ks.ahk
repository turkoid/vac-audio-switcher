#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.

SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
DetectHiddenWindows, on
Menu, Tray, NoStandard ; remove standard Menu items

class VerbosityLevelConfig {
    value := 0
    text := ""
    
    __New(value, text) {
        this.value := value
        this.text := text
    }
}
global log := "vac-ks.log"

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

global MessageBoxVerbosity := VERBOSITY_LEVEL.None
MessageBoxVerbosity := MessageBoxVerbosity | VERBOSITY_LEVEL.Critical

Log(level, message) { 
    FormatTime, timestamp, , % "yyyy-MM-dd HH:mm:ss"
    text :=  timestamp . "`t" . VERBOSITY_LEVEL_CONFIG[level].text . "`t" . message . "`n"
    FileAppend, % text, % log
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

LogImportant(message) {
    Log(VERBOSITY_LEVEL.None, "`n`n" . message . "`n")
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
VAC_ICONS.hdmi := {file: "hdmi-white.ico"}

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
            Menu, tray, icon, % options.file, % options.iconNum
        } else { 
            Menu, tray, icon, % "icons\" . options.file
        }
    }
}

;sets the icon the default one
SetTrayIcon(VAC_ICONS.default)

class VACDevice {
    name := ""
    device := ""
    isVolatile := false
    
    ;constructor
    __New(name, device) {
        this.name := name
        this.device := device
    }
}

;class to store individual settings for each repeater started
;you can override default settings
class VACRepeaterSettings {
    outputDevice := {}
    output := ""
    skip := false
    
    ;constructor
    __New(output) {
        this.outputDevice := output
        this.output := output.device
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
    devices := {}    
    name := ""
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
        this.input := input.device
        this.settings := settings
        this.updateTrayIcon := updateTrayIcon
        this.devices.input := input 
        this.devices.output := settings.outputDevice
        this.name := this.devices.input.name . " -> " . this.devices.output.name
    }
    
    __Get(setting) {
        return this.settings[setting]
    }
    
    IsVolatile() {
        return this.devices.input.isVolatile || this.devices.output.isVolatile
    }
    
    Init() {
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
        if (this.stopParams == "") {
            pb := new ParamBuilder(this)
            pb.Add("closeInstance")
            this.stopParams := pb.params
            pb := {}            
        }         
    }
        
    Open() {
        if (!WinExist(this.winName)) {
            LogInfo("Starting repeater with params: " . this.startParams)            
            target := """" . appPath . """ " . this.startParams
            Run, %target%, , % this.settings.winTarget, pid
            LogInfo("Repeater started: " . this.name)
            this.pid := pid
            WinWait, % this.winName, 1
            if (WinExist("ahk_pid " . pid)) {
                WinGetTitle, windowTitle
                if (windowTitle == "Error") {
                    ControlGetText, errorMessage, Static2, % "ahk_pid " . pid
                    LogError("Unable to open " . this.name ": " . errorMessage)
                    this.Kill()                    
                } else {
                    WinActivate
                    WinHide
                }
            }
        }
    }
        
    Close() {
        if (WinExist(this.winName)) {
            LogInfo("Stopping repeater with params: " . this.stopParams)
            target := """" . appPath . """ " . this.stopParams
            Run, %target%, , Hide
            LogInfo("Repeater stopped: " . this.name)
            WinWaitClose, , , 3
        }
    }
    
    Kill() {
        LogInfo("Repeater killed: " . this.name)
        WinKill, % "ahk_pid " . this.pid    
        SetTrayIcon(VAC_ICONS.default)
    }
    
    Start() {
        if (!WinExist(this.winName)) {
            this.Open()
        }
        if (WinExist(this.winName)) {
            ControlClick, Start, % this.winName
            if (this.updateTrayIcon) {
                SetTrayIcon(this.settings.icon)
            }
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

    Show() {
        if (!WinExist(this.winName)) {
            this.Open()
        }
        if (WinExist(this.winName)) {
            WinShow
        }
    }
    
    Hide() {
        if (WinExist(this.winName)) {
            WinHide
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

class VACRepeaterSet {
    name := ""
    output := {}
    repeaters := {}
    
    __New(output, repeaters) {
        this.output := output
        Loop % repeaters.Length() {
            repeater := repeaters[A_Index]
            this.repeaters.Push(repeater)
            this.name .= repeater.devices.input.Name
            if (A_index != repeaters.Length()) {
                this.name .= "/"
            }
        }
        this.name .= " -> " . output.name
    }  
    
    MaintainSet(op, isCurrentSet := false) {
        menuOp := ""
        Loop % this.repeaters.Length() {
            repeater := this.repeaters[A_Index]
            if (op == "open") {
                repeater.Open()
            } else if (op == "close") {
                repeater.Close()
                menuOp := "stop"
            } else if (op == "kill") {
                repeater.Kill()
                menuOp := "stop"
            } else if (op == "activate") {
                repeater.activate()
            } else if (op == "start") {
                repeater.Start()
                menuOp := "start"
            } else if (op == "stop") {
                repeater.Stop()
                menuOp := "stop"
            } else if (op == "restart") {
                repeater.Restart()
                menuOp := "start"
            } else if (op == "startup") {
                repeater.Init()
                if (!repeater.isVolatile()) {
                    if (isCurrentSet) {
                        repeater.Start()
                        menuOp := "start"
                    } else {
                        repeater.Open()
                    }
                } else {
                    LogInfo("Not starting " . this.name . ": volatile device")
                }
            }
        }    
        if (menuOp == "start") {
            Menu, tray, Check, % this.name
        } else if (menuOp == "stop") {
            Menu, tray, UnCheck, % this.name
        }
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
    isSwitchable := false
    
    __New(name) {
        this.name := name
    }
    
    Init() {        
        if (this.inputs.Length() > 0 && this.outputs.Length() > 0) {         
            Loop % this.outputs.Length() {
                this.isSwitchable |= !this.outputs[outputIndex].skip
                outputIndex := A_Index
                repeaters := {}
                Loop % this.inputs.Length() {
                    inputIndex := A_Index
                    winName := "VAC (" . this.name . "): Input " . inputIndex . " - Output " . outputIndex
                    repeater := new VACRepeater(winName, this.inputs[inputIndex], this.outputs[outputIndex], this.updateTrayIcon)
                    repeaters.Push(repeater)
                }
                set := new VACRepeaterSet(this.outputs[outputIndex].outputDevice, repeaters)                
                this.repeaterSets.Push(set)
            }
        }
        this.setIndex := new WrapAroundIndex(1, this.repeaterSets.Length(), 1)
    }
    
    MaintainAllSets(op) {
        Loop % this.repeaterSets.Length() {
            this.repeaterSets[A_Index].MaintainSet(op)
        }
    }
    
    MaintainCurrentSet(op) {
        this.repeaterSets[this.setIndex.index].MaintainSet(op)
    }
    
    Startup() {
        Loop % this.repeaterSets.Length() {
            this.repeaterSets[A_Index].MaintainSet("startup", A_Index == this.setIndex.index)
        }
    }
    
    Open() {
        this.MaintainAllSets("open")
    }
    
    Close() {
        this.MaintainAllSets("close")
    }
    
    Kill() {
        this.MaintainAllSets("kill")
    }
    
    Start() {
        this.MaintainCurrentSet("start")
        this.paused := false
    }
    
    Stop() {
        this.MaintainCurrentSet("stop")
    }
    
    Restart() {
        this.Stop()
        this.Start()
    }
    
    Activate() {
        this.MaintainAllSets("activate")
    }
    
    Pause() { 
        this.paused := true
        this.Stop()
    }
    
    Resume() {
        this.Start()        
    }
    
    Switch(inc) {
        if (this.paused) {
            this.Resume()
        } else if (!this.persistent) {
            this.Stop()
            if (this.isSwitchable) {
                Loop {       
                    this.setIndex.Increment(inc)
                    repeater := this.repeaterSets[this.setIndex.index].repeaters[1]
                } until (!repeater.skip)                
            } else {
                this.setIndex.Increment(inc)
            }
            this.Start()
        } else {            
            this.Start()
        }
    }
    
    SwitchTo(index) {
        if (!this.persistent) {
            this.Stop()
            this.setIndex.SetIndex(index) 
            this.Start()
        }
    }
}

;NOTE:  For KS, you need use the names in the Audio Repeater KS application, not the control panel sound devices.
;virtual lines created (note: you will get an error if the output device can only support one input).

;stores the repeater setups
global switchers := {}

LoadConfig() {
    LogInfo("LoadConfig - Begin")
    ;Default Device -> Switching Repeater
    ;For me i need this so i can use dxtory to seperate out audio streams
    switcher := new VACSwitcher("Default Device")
    switcher.persistent := true
    device := new VACDevice("Default Device", "Virtual Cable 1")
    switcher.inputs.Push(device)
    device := new VACDevice("Mixer", "Virtual Cable 3")
    settings := new VACRepeaterSettings(device)
    settings.totalBuffer := 70
    switcher.outputs.Push(settings)
    switchers.Push(switcher)

    ;Default comm device -> Switching repeater
    ;need this to capture VOIP stream seperately
    switcher := new VACSwitcher("Default Comm")
    switcher.persistent := true
    device := new VACDevice("Default Comm", "Virtual Cable 2")
    switcher.inputs.Push(device)
    device := new VACDevice("Mixer", "Virtual Cable 3")
    settings := new VACRepeaterSettings(device)
    settings.totalBuffer := 70
    switcher.outputs.Push(settings)
    switchers.Push(switcher)

    ;;;START MIXER

    ;outputs the streams to the correct external output device
    switcher := new VACSwitcher("Mixer")
    switcher.updateTrayIcon := true
    device := new VACDevice("Mixer", "Virtual Cable 3")
    switcher.inputs.Push(device)

    ;speakers
    device := new VACDevice("Speakers", "Sound Blaster Speaker/Headphone")
    settings := new VACRepeaterSettings(device)
    settings.icon := VAC_ICONS.speakers
    switcher.outputs.Push(settings)

    ;headphones
    device := new VACDevice("Headphones", "Sound Blaster SPDIF-Out")
    settings := new VACRepeaterSettings(device)
    settings.icon := VAC_ICONS.headphones
    switcher.outputs.Push(settings)

    ;hdmi
    device := new VACDevice("HDMI", "Unknown")
    device.isVolatile := true
    settings := new VACRepeaterSettings(device)
    settings.icon := VAC_ICONS.hdmi
    settings.skip := true
    switcher.outputs.Push(settings)

    ;Bluetooth speakers
    device := new VACDevice("Bluetooth", "Unknown")
    device.isVolatile := true
    settings := new VACRepeaterSettings(device)
    settings.totalBuffer := 20
    settings.samplingRate := 44100
    settings.icon := VAC_ICONS.bluetooth
    settings.skip := true
    ;switcher.outputs.Push(settings)

    switchers.Push(switcher)
    ;;;END Mixer
    LogInfo("LoadConfig - End")
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
        } else if (op == "kill") {
            switcher.kill()
        } else if (op == "activate") {
            switcher.Activate()
        } else if (op == "startup") {
            switcher.Startup()
        }
    }
}

MenuHandler(itemName, itemPos, menuName) {
    if (menuName == "ShowMenu") {
        Loop % switchers.Length() {
            switcher := switchers[A_Index]
            Loop % switcher.repeaterSets.Length() {
                repeaters := switcher.repeaterSets[A_Index].repeaters
                Loop % repeaters.Length() {
                    repeater := repeaters[A_Index]
                    if (repeater.name == itemName) {
                        repeater.Show()
                    }
                }
            }
        }
    } else {
        if (itemName == "Start") {
            MaintainSwitchers("resume", 0)
        } else if (itemName == "Stop") {
            MaintainSwitchers("pause", 0)
        } else if (itemName == "Exit") {
            MaintainSwitchers("kill", 0)
            ExitApp 0
        } else {            
            Loop % switchers.Length() {
                switcher := switchers[A_Index]
                Loop % switcher.repeaterSets.Length() {
                    set := switcher.repeaterSets[A_Index]
                    if (set.name == itemName) {
                        if (switcher.persistent) {
                            switcher.Start()
                        } else {
                            switcher.switchTo(A_Index)
                        }
                    }
                }                    
            }
        }
    }
}

BuildTrayMenu() {
    LogInfo("BuildingTrayMenu - Begin")
    MaintainSwitchers("init", 0)
    Loop % switchers.Length() {
        switcherIndex := A_Index
        switcher := switchers[switcherIndex]
        
        Loop % switcher.repeaterSets.Length() {        
            setIndex := A_index
            set := switcher.repeaterSets[setIndex]
            Loop % set.repeaters.Length() {
                repeater := set.repeaters[A_Index]
                Menu, ShowMenu, add, % repeater.name, MenuHandler
            }       
            if (switcher.persistent) {
                Menu, tray, add, % set.name, MenuHandler, +Radio    
            } else {
                Menu, tray, add, % set.name, MenuHandler, +Radio
            }
        }    
        if (switcherIndex != switchers.Length()) {
            Menu, ShowMenu, add
        }
        Menu, tray, add
    }
    Menu, tray, add, Show, :ShowMenu
    Menu, tray, add
    Menu, tray, add, Start, MenuHandler
    Menu, tray, add, Stop, MenuHandler
    Menu, tray, add
    Menu, tray, add, Exit, MenuHandler    
    LogInfo("BuildingTrayMenu - End")
}

StartApp() {
    LogImportant("`tAPPLICATION START")
    OnExit("ShutdownApp")
    LoadConfig()
    if (switchers.Length() == 0) {
        LogCritical("No switchers defined!")
        ExitApp 1
    }
    BuildTrayMenu()
    MaintainSwitchers("startup", 0)
}

ShutdownApp() {
    LogImportant("`tAPPLICATION SHUTDOWN")
}

StartApp()

;ctrl+shift+F12
;^+F12::
;    MaintainSwitchers("pause", 0)
;return 

;ctrl+shift+alt+F12
;^+!F12::
;    MaintainSwitchers("close", 0)
;    ExitApp 0
;return

;ctrl+F12
^F12::
    MaintainSwitchers("switch", 1) 
return