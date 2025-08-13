
#Requires AutoHotkey v2.0
; ==============================
; CAC Lookup Add-on for AutoHotkey (AHK v2)
; Hotkey: Win+Alt+C
; Fetches MESA CAC non-zero probability and percentile across sex/race combos.
; Includes "Copy Paragraph" (no table/percentiles) for reports.
; v2-safe: parentheses on IF conditions, one-line hotkey, corrected regex.
; ==============================

; ---- Config ----
global MesaEndpoint := "https://www.mesa-nhlbi.org/Calcium/input.aspx"
global MesaRaces := ["white","black","hispanic","chinese"]
global MesaRaceLabel := Map("white","White","black","Black","hispanic","Hispanic","chinese","Chinese")
global MesaSexes := ["male","female"]
global MesaRequestTimeoutMs := 15000
global MesaInterRequestDelayMs := 250
global AutoCopyParagraphOnLoad := false

; ---- Hotkey (one-liner; no braces) ----
#!c:: CAC_PromptAndShow()

; ---- UI ----
CAC_PromptAndShow() {
    CAC_Gui := 0, CAC_Edit := 0, CAC_Para := 0

    ageIB := InputBox("Enter age (45–84):", "CAC Lookup", "W240 H130")
    if (ageIB.Result = "Cancel")
        return
    age := Trim(ageIB.Value)
    if !CAC_IsInteger(age) {
        MsgBox("Age must be an integer between 45 and 84.", "CAC Lookup", "Icon!")
        return
    }
    age := age + 0
    if (age < 45 || age > 84) {
        MsgBox("The MESA tool supports ages 45–84. You entered " age ".", "CAC Lookup", "Icon!")
        return
    }

    scoreIB := InputBox("Enter Agatston calcium score (≥ 0):", "CAC Lookup", "W280 H140")
    if (scoreIB.Result = "Cancel")
        return
    score := Trim(scoreIB.Value)
    if !CAC_IsNumber(score) || (score+0 < 0) {
        MsgBox("Calcium score must be a non-negative number.", "CAC Lookup", "Icon!")
        return
    }
    score := score + 0

    CAC_Gui := Gui("+AlwaysOnTop +ToolWindow", "CAC Reference (MESA)")
    CAC_Gui.SetFont("s10", "Consolas")
    CAC_Edit := CAC_Gui.Add("Edit", "w760 r18 ReadOnly")
    CAC_Gui.Add("Text", "xm w760 Section", "Paragraph for report (no table / no percentiles):")
    CAC_Para := CAC_Gui.Add("Edit", "w760 r4 ReadOnly")
    btnCopyTable := CAC_Gui.Add("Button", "xs w120 h28", "Copy Table")
    btnCopyPara  := CAC_Gui.Add("Button", "x+8 w140 h28", "Copy Paragraph")
    btnCopyTable.OnEvent("Click", (*) => CAC_Copy(CAC_Edit.Value))
    btnCopyPara.OnEvent("Click", (*) => CAC_Copy(CAC_Para.Value))
    CAC_Gui.Show()

    CAC_Edit.Value := "Fetching from MESA...`r`nThis may take a few seconds."
    results := Mesa_GetAll(age, score)

    text := Mesa_FormatResults(age, score, results)
    CAC_Edit.Value := text

    ptext := Mesa_FormatParagraph(age, score, results)
    CAC_Para.Value := ptext

    if (AutoCopyParagraphOnLoad) {
        A_Clipboard := ptext
        ToolTip("Paragraph copied to clipboard.")
        SetTimer(() => ToolTip(), -1200)
    }
}

CAC_Copy(str) {
    A_Clipboard := str
    ToolTip("Copied to clipboard.")
    SetTimer(() => ToolTip(), -1500)
}

; ---- Fetch loop ----
Mesa_GetAll(age, score) {
    results := Map()
    for _, sex in MesaSexes {
        results[sex] := Map()
        for _, race in MesaRaces {
            r := Mesa_Query(age, sex, race, score)
            results[sex][race] := r
            Sleep(MesaInterRequestDelayMs)
        }
    }
    return results
}

; ---- Single query (IE COM) ----
Mesa_Query(age, sex, race, score) {
    r := Map("prob","", "percentile","", "error","")

    ie := 0
    try {
        ie := ComObject("InternetExplorer.Application")
    } catch as e {
        r["error"] := "COM/IE engine not available"
        return r
    }
    try {
        ie.Visible := false
        ie.Silent := true
        ie.Navigate(MesaEndpoint)
    } catch as e {
        r["error"] := "Navigate failed"
        try ie.Quit()
        return r
    }
    if !Mesa_WaitLoad(ie, MesaRequestTimeoutMs) {
        try ie.Quit()
        r["error"] := "Timeout loading input page"
        return r
    }

    doc := ie.document
    inputs := doc.getElementsByTagName("input")
    count := (inputs ? inputs.length : 0)

    ; --- Robust radio selection ---
    ; gender/sex radios
    Loop count {
        idx := A_Index - 1
        el := inputs.item(idx)
        t := "", n := "", v := "", id := "", html := ""
        try t := StrLower(el.type)
        if (t != "radio")
            continue
        try n := StrLower(el.name)
        try v := StrLower(el.value)
        try id := StrLower(el.id)
        try html := StrLower(el.outerHTML)
        if !(InStr(n,"gender") || InStr(n,"sex"))
            continue
        if (sex="male" && (InStr(v,"male") || InStr(id,"male") || InStr(html, ">male<")))
            el.checked := true
        if (sex="female" && (InStr(v,"female") || InStr(id,"female") || InStr(html, ">female<")))
            el.checked := true
    }
    ; race/ethnicity radios
    Loop count {
        idx := A_Index - 1
        el := inputs.item(idx)
        t := "", n := "", v := "", id := "", html := ""
        try t := StrLower(el.type)
        if (t != "radio")
            continue
        try n := StrLower(el.name)
        try v := StrLower(el.value)
        try id := StrLower(el.id)
        try html := StrLower(el.outerHTML)
        if !(InStr(n,"race") || InStr(n,"ethnic"))
            continue
        if (race="white" && (InStr(v,"white") || InStr(id,"white") || InStr(html, "white")))
            el.checked := true
        if (race="black" && (InStr(v,"black") || InStr(id,"black") || InStr(html, "african") || InStr(html,"black")))
            el.checked := true
        if (race="hispanic" && (InStr(v,"hispanic") || InStr(id,"hispanic") || InStr(html, "hispanic")))
            el.checked := true
        if (race="chinese" && (InStr(v,"chinese") || InStr(id,"chinese") || InStr(html, "chinese")))
            el.checked := true
    }
    ; --- Age & Score ---
    Loop count {
        idx := A_Index - 1
        el := inputs.item(idx)
        t := "", n := ""
        try t := StrLower(el.type)
        if (t != "text")
            continue
        try n := StrLower(el.name)
        if (InStr(n,"age"))
            el.value := age
        else if (InStr(n,"score") || InStr(n,"agatston"))
            el.value := score
    }

    ; --- Submit ---
    btn := 0
    Loop count {
        idx := A_Index - 1
        el := inputs.item(idx)
        t := "", val := "", id := "", name := ""
        try t := StrLower(el.type)
        try val := StrLower(el.value), id := StrLower(el.id), name := StrLower(el.name)
        if (t="submit" && (InStr(val,"calculate") || InStr(id,"calc") || InStr(name,"calc"))) {
            btn := el
            break
        }
    }
    if (!btn) {
        buttons := doc.getElementsByTagName("button")
        bcount := (buttons ? buttons.length : 0)
        Loop bcount {
            idx := A_Index - 1
            el := buttons.item(idx)
            inner := ""
            try inner := StrLower(el.innerText)
            if (InStr(inner,"calculate")) {
                btn := el
                break
            }
        }
    }
    if (btn) {
        btn.click()
    } else if (doc.forms.length > 0) {
        try doc.forms[0].submit()
    }

    ; --- Wait for results ---
    Mesa_WaitReadyText(ie, "percentile", MesaRequestTimeoutMs)
    Sleep(150)

    txt := ""
    try txt := ie.document.body.innerText
    try ie.Quit()

    if (txt = "") {
        r["error"] := "Empty response"
        return r
    }

    ; --- Parse probability and estimated percentile ---
    prob := Mesa_RegexNum(txt, "(?is)probability\s+of\s+(?:a\s+)?non[-\s]?zero[^%]*?(\d{1,3}(?:\.\d+)?)\s*%")
    if (prob = "")
        prob := Mesa_RegexNum(txt, "(?is)non[-\s]?zero[^%]*?probab[^%]*?(\d{1,3}(?:\.\d+)?)\s*%")

    pct := Mesa_RegexNum(txt, "(?is)(?:estimated\s+percentile(?:\s+for\s+this\s+score)?|your\s+score\s+is\s+at\s+the)\D*?(\d{1,3})(?:st|nd|rd|th)?\s*percentile")
    if (pct = "")
        pct := Mesa_RegexNum(txt, "(?is)\bat\s+the\s+(\d{1,3})(?:st|nd|rd|th)?\s*percentile")

    r["prob"] := prob
    r["percentile"] := pct
    return r
}

; ---- Helpers ----
Mesa_WaitLoad(ie, timeout := 15000) {
    t0 := A_TickCount
    while (ie.busy || ie.ReadyState != 4) {
        Sleep(50)
        if (A_TickCount - t0 > timeout)
            return false
    }
    return true
}

Mesa_WaitReadyText(ie, needle, timeout := 15000) {
    t0 := A_TickCount
    ndl := StrLower(needle)
    loop {
        txt := ""
        try txt := ie.document.body.innerText
        if (StrLen(txt) && InStr(StrLower(txt), ndl))
            return true
        if (A_TickCount - t0 > timeout)
            break
        Sleep(100)
    }
    return false
}

Mesa_RegexNum(haystack, pattern) {
    m := []
    if RegExMatch(haystack, pattern, &m)
        return m[1]
    return ""
}

Mesa_FormatResults(age, score, results) {
    out := ""
    out .= "Age: " age "    Score: " score "`r`n"
    out .= "MESA CAC Reference — non-zero probability & score percentile`r`n"
    out .= "-----------------------------------------------------------------------`r`n"
    for _, sex in MesaSexes {
        sexLabel := (sex = "male") ? "Male" : "Female"
        out .= "[" sexLabel "]`r`n"
        for _, race in MesaRaces {
            label := MesaRaceLabel[race]
            r := results[sex][race]
            if r.Has("error") && r["error"] != "" {
                line := Format("{:10}:  error: {}", label, r["error"])
            } else {
                p := (r["prob"] != "" ? r["prob"] . "%" : "n/a")
                pct := (r["percentile"] != "" ? r["percentile"] : "n/a")
                line := Format("{:10}:  non-zero {:>6}    percentile {:>5}", label, p, pct)
            }
            out .= "  " line "`r`n"
        }
        out .= "`r`n"
    }
    out .= "Note: Values pulled live from MESA (" MesaEndpoint "). Internet required.`r`n"
    return out
}

Mesa_FormatParagraph(age, score, results) {
    male := results["male"], female := results["female"]
    getp := (m) => (m && m.Has("prob") && m["prob"] != "" ? m["prob"] . "%" : "n/a")

    m_white := getp(male["white"]),   m_black := getp(male["black"]),   m_hisp := getp(male["hispanic"]),   m_chin := getp(male["chinese"])
    f_white := getp(female["white"]), f_black := getp(female["black"]), f_hisp := getp(female["hispanic"]), f_chin := getp(female["chinese"])

    para := "MESA CAC (age " age ", Agatston " score "): non-zero CAC probability — "
    para .= "Male [White " m_white ", Black " m_black ", Hispanic " m_hisp ", Chinese " m_chin "]; "
    para .= "Female [White " f_white ", Black " f_black ", Hispanic " f_hisp ", Chinese " f_chin "]."
    return para
}

CAC_IsInteger(val) {
    return RegExMatch(val, "^[+-]?\d+$")
}
CAC_IsNumber(val) {
    return RegExMatch(val, "^[+-]?\d+(\.\d+)?$")
}