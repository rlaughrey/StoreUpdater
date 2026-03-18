#Requires AutoHotkey v2.0
#include OCR.ahk
#SingleInstance

updateMSStore() {
    if !WinExist("Microsoft Store") {
        Run "ms-windows-store://updates"
        Sleep 7000
    } else {
        WinActivate "Microsoft Store"
        Sleep 5000
    }

    ;; OCR the window
    result := OCR.FromWindow("Microsoft Store")

    if !result {
        MsgBox "OCR failed"
        return
    }

    ;; Find and click "Check for updates"
    for line in result.Lines {
        text := Trim(line.Text)

        if (InStr(text, "Check for updates")) {
            x := line.X + (line.W // 2)
            y := line.Y + (line.H // 2)

            Click x, y
            Sleep 8000 ; wait for scan to finish
            break
        }
    }

    ;; Re-scan for "Update all" (UI changes after check)
    result := OCR.FromWindow("Microsoft Store")

    for line in result.Lines {
        text := Trim(line.Text)

        if (InStr(text, "Update all")) {
            x := line.X + (line.W // 2)
            y := line.Y + (line.H // 2)

            Click x, y
            break
        }
    }
}

Run "ms-windows-store://updates"
Sleep 7000

SetTimer(updateMSStore, 10000)

Sleep 900000
ExitApp