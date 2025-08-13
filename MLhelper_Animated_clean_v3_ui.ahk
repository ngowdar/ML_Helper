#Include Acc.ahk
#Include UIA.ahk
#Include ColorButton.ahk
#Include %A_ScriptDir%\CAC_Lookup_addon_v2_paragraph_scoped_fix6.ahk

global CONFIG_FILE := A_ScriptDir "\user_input_defaults.ini"

global studyCounter := 0
global CTCounter := 0
global USCounter := 0
global MRCounter := 0
global CTweight := 1.0
global USweight := 0.333
global MRweight := 1.0
global weightedcount := 0
global pathString := ""
global runningText := ""
global trans := false
global ProviderFullName := ""
global provLastName := ""
global provFirstName := ""
global PatientFullName := ""
global PatientLastName := ""
global PatientFirstName := ""
global PatientMRN := ""
global SendersText := ""
global PageText := ""
global ListBoxShowing := true
global SenderNameEdit := ""
global PageTextEdit := ""
global AnimationsEnabled := true
global ShowTallyAfterLog := true
global TallyDisplayDuration := 3000
global CurrentAccNumber := ""
global LastAccNumber := ""
global StudyType := ""
global unpagedcolor := "17dc97"
global pagedcolor := "939393"
global unloggedcolor := "2490da"
global loggedcolor := "939393"
global blacktext := "000000"
global whitetext := "FFFFFF"
filenameTxt := ""
date := FormatTime(, "MM-dd-yyyy")
accessionNumber := ""
global lastStudyType := ""
ordProvider := ""
curAcc := ""
global StudyLog := []
duration := 700

; Create the GUI
myGui := Gui("+AlwaysOnTop -Caption -ToolWindow -Border -Resize -MaximizeBox +LastFound")
myGui.BackColor := "215f6e"
WinSetTransparent(200)
myGui.SetFont("s10 cWhite")
LogBtn := myGui.Add("Button", "w175 h28", "LOG")
myGui.Add("Text", "ys w125 cWhite vLastAcc", "Last Logged Acc: 00000000")
PageBtn := myGui.Add("Button", "ys w150 h28", "PAGE")
;myGui.Add("Text", "w100 cYellow vCurrentAcc", "Current Acc: Xx-N/A-xX")
;myGui.Add("Text", "ys w130 cYellow vPatientInfo", "Current Patient: Doe, JOHN N/A")
;myGui.Add("Text", "ys w100 cAqua vProvider", "Ord. Prov: Doe, John N/A")
myGui.Add("Text", "ys w210 cYellow vCounterText", "Total Count: 0  | CT: 0  | US: 0  | MRI: 0")
myGui.Add("Text", "ys w125 cAqua vWeightedCountText", "Weighted Count: 0")
ShowLogBtn := myGui.Add("Button", "ys w150 h28", "Show List").OnEvent("Click", ShowLog)
LogBtn.OnEvent("Click", LogButton)
PageBtn.OnEvent("Click", PageButton)

;LogBtn.SetColor(unloggedcolor, blacktext)
;PageBtn.SetColor(unpagedcolor, blacktext)
LogBtn.TextColor := blacktext
LogBtn.backColor := unloggedcolor
PageBtn.TextColor := blacktext
PageBtn.backColor := unpagedcolor
GUIcenterX := ((A_ScreenWidth)*0.5)+275
GUIcenterY := 30
;myGui.Show("x" GUIcenterX-350 " y" 1 " AutoSize")
myGui.Show("x" GUIcenterX " y" 1)

; Tab 2
;tab.UseTab(2)
;myGui.SetFont("s8 cRed")
;listBox := myGui.Add("ListBox", "r8 vscroll vListSelect")
;myGui.Add("Button", "w120", "Log Current").OnEvent("Click", LogButton)
;myGui.Add("Button", "w120", "Delete").OnEvent("Click", RemoveSelectedEntry)
listXposition := ((A_ScreenWidth)*0.5)+280
listGui := Gui("+AlwaysOnTop -Caption -ToolWindow -Border -Resize -MaximizeBox +LastFound")
listGui.BackColor := "215f6e"
WinSetTransparent(200)
listGui.SetFont("s10 cWhite")
listBox := listGui.Add("ListBox", "r8 vscroll vListSelect")
listGui.Add("Button", "w120 h24", "Log Current").OnEvent("Click", LogButton)
listGui.Add("Button", "w120 h24", "Delete").OnEvent("Click", RemoveSelectedEntry)
listGui.Show("x" listXposition " y" 40 " AutoSize Hide")

PositionGUI()

; Create intro option box
firstGui := Gui("-Border -Caption")
firstGui.BackColor := "215f6e"
firstGui.SetFont("s10 cWhite")
firstGui.Add("Button", "w150", "Select Folder to Save Log").OnEvent("Click", SelectSaveFolder)
firstGui.Add("Button", "w150", "Append Existing Log").OnEvent("Click", ExistingFile)
firstGui.Add("Button", "w150", "Save as Custom Filename").OnEvent("Click", CustomFilename)
firstGui.Add("Text", "w150 Center cWhite", "Animation Settings:")
animToggle := firstGui.Add("Checkbox", "w150 Checked" (AnimationsEnabled ? "1" : "0"), "Enable Animations")
animToggle.OnEvent("Click", (*) => ToggleAnimations(animToggle))
firstGui.Show()
AW_BLEND := 0x00080000
AW_HIDE := 0x00010000
AW_SLIDE := 0x00040000
AW_HOR_NEGATIVE := 0x00000002
AW_HOR_POSITIVE := 0x00000001 

SetTimer(CheckForNewReport, 3000)

PositionGUI() { 
	
	GUIcenterX := (A_ScreenWidth)*0.5
	GUIcenterY := 30
    myGui.Show("x" GUIcenterX-350 " y" 1 " AutoSize")
	;AnimateWindow(myGui.Hwnd, 250, AW_BLEND)
}

ToggleAnimations(checkboxObj) {
    global AnimationsEnabled, ShowTallyAfterLog
    AnimationsEnabled := checkboxObj.Value
    ShowTallyAfterLog := checkboxObj.Value
}

#!g::
{
	AW_BLEND := 0x00080000
	AW_HIDE := 0x00010000
	AW_SLIDE := 0x00040000
	AW_HOR_NEGATIVE := 0x00000002
	AW_HOR_POSITIVE := 0x00000001
	;listGui.Show()
	AnimateWindow(listGui.Hwnd, 250, AW_BLEND)	
	;ShowLog()
}

RemoveSelectedEntry(*) {
    global StudyLog, CTCounter, USCounter, MRCounter, studyCounter
	selected := listBox.Value
    if (selected) {
		templine := StudyLog[selected]
		if InStr(A_LoopField, "US ") {
			USCounter--
			studyCounter--			
		}
		if InStr(A_LoopField, "CT ") {
			CTCounter--
			studyCounter--			
		}
		if InStr(A_LoopField, "MR ") {
			MRCounter--
			studyCounter--			
		}
		StudyLog.RemoveAt(selected)
		listBox.Delete(selected)
		UpdateStudyLog()
	}     	
}

AnimateWindow(hWin, Time, Flags) {
    return DllCall("User32.dll\AnimateWindow", "Ptr", hWin, "UInt", Time, "UInt", Flags)
}

SlideInFromRight(guiObj, duration := 300) {
    global AnimationsEnabled
    if (!AnimationsEnabled) {
        guiObj.Show("AutoSize")
        return
    }
    
    ; Get screen dimensions and calculate slide-in position
    screenWidth := A_ScreenWidth
    guiObj.Show("Hide AutoSize")
    guiObj.GetPos(,, &width, &height)
    
    ; Start from off-screen right
    startX := screenWidth
    endX := screenWidth - width - 20
    startY := 50
    
    ; Show window at starting position
    guiObj.Show("x" startX " y" startY " AutoSize NoActivate")
    
    ; Smooth slide animation
    steps := 20
    stepDelay := duration // steps
    stepDistance := (startX - endX) // steps
    
    Loop steps {
        currentX := startX - (stepDistance * A_Index)
        guiObj.Move(currentX, startY)
        Sleep(stepDelay)
    }
    
    ; Ensure final position is exact
    guiObj.Move(endX, startY)
}

SlideOutToRight(guiObj, duration := 300) {
    global AnimationsEnabled
    if (!AnimationsEnabled) {
        guiObj.Hide()
        return
    }
    
    ; Get current position and screen width
    guiObj.GetPos(&currentX, &currentY, &width)
    screenWidth := A_ScreenWidth
    endX := screenWidth
    
    ; Smooth slide animation
    steps := 20
    stepDelay := duration // steps
    stepDistance := (endX - currentX) // steps
    
    Loop steps {
        newX := currentX + (stepDistance * A_Index)
        guiObj.Move(newX, currentY)
        Sleep(stepDelay)
    }
    
    ; Hide the window
    guiObj.Hide()
}

ShowTallyTemporarily() {
    global myGui, TallyDisplayDuration, ShowTallyAfterLog
    if (!ShowTallyAfterLog)
        return
        
    ; Create a temporary tally display GUI
    tallyGui := Gui("+AlwaysOnTop -Caption -ToolWindow -Border +LastFound")
    tallyGui.BackColor := "2d5a3d"
    WinSetTransparent(220, tallyGui)
    tallyGui.SetFont("s12 cWhite Bold")
    
    ; Get current counts for display
    global studyCounter, CTCounter, USCounter, MRCounter, weightedcount
    tallyText := "Study Logged!`n"
    tallyText .= "Total: " studyCounter " | CT: " CTCounter " | US: " USCounter " | MRI: " MRCounter "`n"
    tallyText .= "Weighted Count: " Round(weightedcount, 2)
    
    tallyGui.Add("Text", "w300 h80 Center", tallyText)
    
    ; Slide in from right
    SlideInFromRight(tallyGui, 250)
    
    ; Auto-hide after duration
    SetTimer(HideTallyGui.Bind(tallyGui), -TallyDisplayDuration)
}

HideTallyGui(tallyGui) {
    SlideOutToRight(tallyGui, 250)
    SetTimer(DestroyTallyGui.Bind(tallyGui), -300)
}

DestroyTallyGui(tallyGui) {
    tallyGui.Destroy()
}

IsAccessionAlreadyLogged(accessionNumber) {
    global StudyLog
    for entry in StudyLog {
        if InStr(entry, accessionNumber) = 1 {
            return true
        }
    }
    return false
}

CheckForNewReport(*) {
    global CurrentAccNumber, LastAccNumber, weightedcount, CTCounter, USCounter, MRCounter, CTweight, USweight, MRweight
	provFirstInit := ""
	if WinExist("Fluency for Imaging Reporting") {
        try {
            fluency := Acc.ElementFromHandle("Fluency for Imaging Reporting")
            CurrentAccNumber := fluency[4,1,1,1,1,3,1,1,2,1,1,1,1,6].Name
            studyType := fluency[4,1,1,1,1,3,1,1,4,2,3,2,1,1].Name
            provWholeName := fluency[4,1,1,1,1,3,1,1,2,1,1,1,2,6].Name
			PatientFullName := fluency[4,1,1,1,1,3,1,1,2,1,1,1,1,1].Name						
			if (CurrentAccNumber != "") {
				ProviderNameArray := StrSplit(provWholeName, ", ")
				ProviderLastName := ProviderNameArray[1]
				ProviderFirstName := ProviderNameArray[2]
				provFirstInit := SubStr(ProviderFirstName, 1, 1)
				PatientNameArray := StrSplit(PatientFullName, ", ")
				PatientLastName := PatientNameArray[1]
				PatientFirstName := PatientNameArray[2]
				PatientFirstInit := SubStr(PatientFirstName, 1, 1)
				ordProvider := ProviderLastName ", " provFirstInit
				;myGui["Provider"].Value := "Provider: " ordProvider
				;myGui["PatientInfo"].Value := "Current Patient: " PatientLastName ", " PatientFirstInit
				;myGui["CurrentAcc"].Value := "Current Acc: " CurrentAccNumber
				LogBtn.Text := "Log: " PatientLastName ", " PatientFirstInit " (" CurrentAccNumber ")"
				PageBtn.Text := "Page: " ordProvider
				weightedcount := (CTweight*CTCounter) + (USweight*USCounter) + (MRweight*MRCounter)
				myGui["WeightedCountText"].Value := "Weighted Count: " weightedcount
				if IsAccessionAlreadyLogged(CurrentAccNumber)
					LogBtn.SetColor(loggedcolor, whitetext)
				else
					LogBtn.SetColor(unloggedcolor, blacktext)
				return
			}
		}
	}   
}

SelectSaveFolder(*) {
    global pathString, filenameTxt, CTCounter, USCounter, MRCounter, studyCounter
	firstGui.Submit()	
    if (folder := DirSelect(, 3)) {
        MsgBox("You selected folder " folder)
        filenameTxt := date " Log.txt"
		pathString := folder "\" date " Log.txt"
    } else {
        pathString := "H:\Moonlighting Logs\" date " Log.txt"
    }
    firstGui.Destroy()
    
    ;guiText := Format("Total Studies: {1}`nCT: {2}`nUS: {3}`nMR: {4}", studyCounter, CTCounter, USCounter, MRCounter)
    ;myGui["PathText"].Value := filenameTxt
    ;myGui["CounterText"].Value := "Total Studies: " studyCounter
	myGui["CounterText"].Value := "Total Count: " studyCounter " | CT: " CTCounter " | US: " USCounter " | MRI: " MRCounter
}

CustomFilename(*) {
    global pathString, filenameTxt, CTCounter, USCounter, MRCounter, studyCounter
    firstGui.Submit()
    
    ; Prompt user for custom filename
    customName := InputBox("Enter custom filename (without .txt extension):", "Custom Log Filename", "w300 h150")
    
    if customName.Result = "Cancel" {
        ExitApp()
    }
    
    if (customName.Value = "") {
        MsgBox("Filename cannot be empty. Using default name.")
        filenameTxt := date " Log.txt"
        pathString := A_ScriptDir "\" filenameTxt
    } else {
        ; Ensure filename has .txt extension
        if !InStr(customName.Value, ".txt")
            filenameTxt := customName.Value ".txt"
        else
            filenameTxt := customName.Value
            
        ; Allow user to select folder for the custom file
        if (folder := DirSelect(, 3, "Select folder to save '" filenameTxt "'")) {
            pathString := folder "\" filenameTxt
        } else {
            pathString := A_ScriptDir "\" filenameTxt
        }
    }
    
    firstGui.Destroy()
    myGui["CounterText"].Value := "Total Count: " studyCounter " | CT: " CTCounter " | US: " USCounter " | MRI: " MRCounter
}

PageButton(*) {    
	global provLastName, provFirstName
	SendersText := "Neel Gowdar, MD (radiology)"
	if WinExist("Fluency for Imaging Reporting") {
        WinActivate("Fluency for Imaging Reporting")
		try {
			mshtaEl := UIA.ElementFromHandle("Fluency for Imaging Reporting ahk_exe mshta.exe")
			PatientFullName := mshtaEl.ElementFromPath("YY/YYYrYbUqbUbUK").Name
			PatientMRN := mshtaEl.ElementFromPath("YY/YYYrYbUqbUbUKs").Name
			PatientAccNum := mshtaEl.ElementFromPath("YY/YYYrYbUqbUbUKu").Name
			ProviderFullName := mshtaEl.ElementFromPath("YY/YYYrYbUqbUbU/Ku").Name
			NameArray := StrSplit(ProviderFullName, ", ")
            provLastName := NameArray[1]
            first := NameArray[2]
            provFirstName := SubStr(first, 1, 1)
			PageText := "Hi, can you call 73707 regarding imaging results for pt " PatientFullName " (" PatientMRN ")? Thanks!"			
		}
		WinMinimize("Fluency for Imaging Reporting")
	}
	
	;Prompt user for text input and to check if default text is okay or needs to be edited first
	SenderDefault := IniRead(CONFIG_FILE, "Defaults", "SenderName", "Default Sender Name")
	PageTextDefault := IniRead(CONFIG_FILE, "Defaults", "PageText", "Default Page Text")
	InputPrompt := Gui()
	InputPrompt.Title := "User Input"
	InputPrompt.Opt("+Resize")
	InputPrompt.Add("Text", "x10 y10 w200", "Sender's Name:")
	SenderNameEdit := InputPrompt.Add("Edit", "x10 y30 w300", SenderDefault)
	InputPrompt.Add("Text", "x10 y60 w200", "Page Text:")
	PageTextEdit := InputPrompt.Add("Edit", "x10 y80 w300 h100", PageTextDefault)
	SaveBtn := InputPrompt.Add("Button", "x10 y190 w100", "Save")
	SaveBtn.OnEvent("Click", SaveHandler)
	SubmitBtn := InputPrompt.Add("Button", "x120 y190 w100", "Submit")
	SubmitBtn.OnEvent("Click", SubmitHandler)
	CancelBtn := InputPrompt.Add("Button", "x230 y190 w100", "Cancel")
	CancelBtn.OnEvent("Click", (*) => ExitApp())
	InputPrompt.OnEvent('Close', (*) => ExitApp())
	InputPrompt.Show()
	;SenderDefault := SendersText
	;PageTextDefault := PageText
	;InputPrompt.Destroy()
	
	PagingWebsite := Gui(, "Paging Window")
	PagingWebsite.Opt("+AlwaysOnTop")
	WB := PagingWebsite.Add("ActiveX", "w700 h460", "Shell.Explorer").Value
	WB.Navigate("http://webpage.rush.edu/smartweb/pages/paging/paging.jsf")
	PagingWebsite.Show()
	sleep(1000)
	;WinActivate("Paging Window")
	if WinExist("Paging Window") {
        WinActivate("Paging Window")
		try {
			AutoHotkeyEl := UIA.ElementFromHandle("Paging Window ahk_exe AutoHotkey32.exe")	
			AutoHotkeyEl.ElementFromPath("YYYYYbrUbqUbUbUq4").Value := provLastName
			AutoHotkeyEl.ElementFromPath("YYYYYbrUbqUbUbUs4").Value := provFirstName
			AutoHotkeyEl.ElementFromPath("YYYYYbrUbqUbU/bU56").Click()
			sleep(300)
			AutoHotkeyEl.ElementFromPath("YYYYYbrUbsUbUr5").Click()
			AutoHotkeyEl.ElementFromPath("YYYYYbuUbUq4").Value := SendersText
			AutoHotkeyEl.ElementFromPath("YYYYYbuUbUs4").Value := PageText
		}		
		;SenderPrompt := InputBox("Provider found. Please review/edit message if needed.", "Enter message",,PageText)
		/* if SenderPrompt.Result = "Cancel"
			MsgBox "Page cancelled."
		else {
			WinActivate("Paging Window")
			PageText2 := SenderPrompt.Value
			AutoHotkeyEl.ElementFromPath("YYYYYbuUbUq4").Value := SendersText
			AutoHotkeyEl.ElementFromPath("YYYYYbuUbUs4").Value := PageText2
			;AutoHotkeyEl.ElementFromPath("YYYYYb/U5").Click()
			PagingWebsite.Destroy()
			MsgBox("Page sent")
			}		
		sleep(300)				
		if WinExist("Paging Window")
			PagingWebsite.Destroy()	 */
	}
	WinActivate("Fluency for Imaging Reporting")
}

LogButton(*) {
    LogPatient()
}

ShowLog(*)  {
	global ListBoxShowing
	if (ListBoxShowing) {
		listGui.Hide()
		ListBoxShowing := false
	}
	else  {
		listXposition := ((A_ScreenWidth)*0.5)+275
		listGui.Show("x" listXposition " y" 32 " AutoSize")
		ListBoxShowing := true
	}
	;MsgBox("Did it show?")
}

DisplayStudyLog() {
	global StudyLog, runningText
	tempText := ""
	Loop StudyLog.Length {
		tempText .= StudyLog[A_Index] "`n"
	}
	MsgBox("StudyLog: " tempText "`nRunningText: " runningText)
}

ExistingFile(*) {
    global CTCounter, USCounter, MRCounter, studyCounter, StudyLog, pathString, LastAccNumber, runningText
	firstGui.Submit()
	/* CTCounter := 0
	USCounter := 0
	studyCounter := 0 */
	pathString := ""
    if (file := FileSelect(3,, "Select an existing log file (*.txt)")) {
        pathString := file
        SplitPath(file,, &filenameTxt)
        fileContent := FileRead(file)
        firstLine := true
        
        Loop Parse, fileContent, "`n", "`r" {
            if (firstLine) {
                if InStr(A_LoopField, "Total Studies: ") {
                    firstLine := false
                    continue
                }
                /* if InStr(A_LoopField, "US ") {
                    USCounter++
                    studyCounter++
                    runningText .= A_LoopField "`n"
                    StudyLog.Push(A_LoopField)
                }
                if InStr(A_LoopField, "CT ") {
                    CTCounter++
                    studyCounter++
                    runningText .= A_LoopField "`n"
                    StudyLog.Push(A_LoopField)
                } */
                listBox.Add([A_LoopField])
                firstLine := false
            } 
			else {
                if (StrLen(A_LoopField) < 8) {
                    break                    
				}
                if InStr(A_LoopField, "US ") {
                    USCounter++
                }
                if InStr(A_LoopField, "CT ") {
                    CTCounter++
                }
				if InStr(A_LoopField, "MR ") {
                    MRCounter++
                }
				studyCounter++
				listBox.Add([A_LoopField])
				runningText .= A_LoopField "`n"
				StudyLog.Push(A_LoopField)
				LastAccNumber := SubStr(A_LoopField, 1, 8)
            }
        }
    }
    
    firstGui.Destroy()    
    ;myGui["CounterText"].Value := "Total Studies: " studyCounter    
	myGui["CounterText"].Value := "Total Count: " studyCounter " | CT: " CTCounter " | US: " USCounter " | MRI: " MRCounter
    myGui["LastAcc"].Value := "Last Logged Acc: " LastAccNumber
	weightedcount := (CTweight*CTCounter) + (USweight*USCounter) + (MRweight*MRCounter)
	myGui["WeightedCountText"].Value := "Weighted Count: " weightedcount
    finalContent := Format("Total Studies: {1}, CT: {2}, US: {3}, MR: {4}`n{5}", studyCounter, CTCounter, USCounter, MRCounter, runningText)
	DisplayStudyLog()
    try {
		file1 := FileOpen(pathString, "w")
        file1.Write(finalContent)
    } catch Error as err {
        MsgBox("Can't open " pathString " for writing.")
    }
}

UpdateStudyLog() {
	global runningText, StudyLog, pathString
	;StudyLog is updated, need to transfer to runningText
	runningText := ""
	Loop StudyLog.Length {
		runningText .= StudyLog[A_Index] "`n"
	}
	/* finalContent := Format("Total Studies: {1}, CT: {2}, US: {3}`n{4}", studyCounter, CTCounter, USCounter, runningText)
	try {
		file1 := FileOpen(pathString, "w")
        file1.Write(finalContent)
    } catch Error as err {
        MsgBox("Can't open " pathString " for writing.")
    } */
	UpdateCountHeader()
}

LogPatient() {
    global CurrentAccNumber, LastAccNumber, lastStudyType, studyCounter, runningText, StudyLog, CTCounter, USCounter, MRCounter
	if WinExist("Fluency for Imaging Reporting") {
		WinActivate("Fluency for Imaging Reporting")
        try {            
            fluency := Acc.ElementFromHandle("Fluency for Imaging Reporting")
            CurrentAccNumber := fluency[4,1,1,1,1,3,1,1,2,1,1,1,1,6].Name
            studyType := fluency[4,1,1,1,1,3,1,1,4,2,3,2,1,1].Name  
			}
		catch {
			ToolTip("Couldn't get Accession Number!")
			SetTimer(() => ToolTip(), -3000)
			return
			}
	}
	if IsAccessionAlreadyLogged(CurrentAccNumber) {
		ToolTip("This has already been logged! " CurrentAccNumber "`nTotal Count So Far is: " studyCounter)
		SetTimer(() => ToolTip(), -3000)
		return
	}			
	else  {		
		LastAccNumber := CurrentAccNumber	
		lastStudyType := studyType
		studyCounter++
		myGui["LastAcc"].Value := "Last Logged Acc: " LastAccNumber					
		textLine := CurrentAccNumber ", " studyType							
		StudyLog.Push(textLine)					
		if InStr(studyType, "US ") {			
			USCounter++
		}
		if InStr(studyType, "CT ") {			
			CTCounter++
		}
		if InStr(studyType, "MR ") {			
			MRCounter++
		}			
		;count := CTCount + (USCount * 0.333)
		runningText .= textLine "`n"					
		UpdateCountHeader()					
		;guiText := Format("Total Studies: {1}`nCT: {2}     US: {3}", studyCounter, CTCounter, USCounter)
		myGui["CounterText"].Value := "Total Count: " studyCounter " | CT: " CTCounter " | US: " USCounter " | MRI: " MRCounter
		;myGui["PathText"].Value := pathString
		listBox.Add([textLine])
		LogBtn.TextColor := blacktext
		LogBtn.backColor := unloggedcolor
		toolText := Format("Logged Accession Number: {1}`nCT: {2} US: {3} MR: {4} Overall: {5}", 
			textLine, CTCounter, USCounter, MRCounter, studyCounter)
		ToolTip(toolText)
		SetTimer(() => ToolTip(), -3000)
		
		; Show animated tally display
		ShowTallyTemporarily()
	}		           
}

;HotIfWinActive("A")
#!l::
{	
	LogPatient()
}

; Toggle animation hotkey
#!a::
{
    global AnimationsEnabled, ShowTallyAfterLog
    AnimationsEnabled := !AnimationsEnabled
    ShowTallyAfterLog := !ShowTallyAfterLog
    
    status := AnimationsEnabled ? "Enabled" : "Disabled"
    ToolTip("Animations " status)
    SetTimer(() => ToolTip(), -2000)
}

/* Hotkey("!p", (*) => {
    if WinExist("Fluency for Imaging Reporting") {
        WinActivate()
        fluency := Acc.ElementFromHandle("Fluency for Imaging Reporting")
        wholeName := fluency[4,1,1,1,3,1,1,3,1,1,1,3,4].Name
        nameArray := StrSplit(wholeName, ", ")
        lastName := nameArray[1]
        first := nameArray[2]
        firstName := SubStr(first, 1, 1)
        
        TrayTip(lastName ", " firstName, "Attempting to page")
        SetTimer(() => {
            TrayTip()
            if SubStr(A_OSVersion, 1, 3) = "10." {
                A_TrayMenu.ClickCount := 0
                Sleep(200)
                A_TrayMenu.ClickCount := 1
            }
        }, -3000)
        
        wb := ComObject("InternetExplorer.Application")
        wb.Navigate("http://webpage.rush.edu/smartweb/pages/paging/paging.jsf")
        wb.Visible := true
        Sleep(1000)
        SendText(lastName)
        Send("{Tab}")
        SendText(firstName)
        Send("{Enter}")
    }
}) */

UpdateCountHeader() {
	global pathString, runningText, studyCounter, CTCounter, USCounter, MRCounter
	file1 := FileOpen(pathString, "w")
    finalContent := Format("Total Studies: {1}, CT: {2}, US: {3}, MR: {4}`n{5}", 
        studyCounter, CTCounter, USCounter, MRCounter, runningText)
    try {
        file1.Write(finalContent)
    } catch Error as err {
        MsgBox("Can't open " pathString " for writing.")
    }
}

SaveHandler(*)
{
    global SendersText, PageText, CONFIG_FILE, SenderNameEdit, PageTextEdit
	if WinExist("User Input") {
		try {
			SendersText := SenderNameEdit.Text
			PageText := PageTextEdit.Text
			IniWrite(SendersText, CONFIG_FILE, "Defaults", "SenderName")
			IniWrite(PageText, CONFIG_FILE, "Defaults", "PageText")    
			MsgBox("Defaults Saved!")
		}
		catch {
			MsgBox("Couldn't save text!")
		}
	}
}

SubmitHandler(*)
{
    global SendersText, PageText, SenderNameEdit, PageTextEdit
    try {
		SendersText := SenderNameEdit.Text
		PageText := PageTextEdit.Text
		MsgBox("Sender: " SendersText "`nPage Text: " PageText)
	}
	catch {
		MsgBox("Couldn't submit page!")
	}    
}