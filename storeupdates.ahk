#Requires AutoHotkey v2.0
#include OCR.ahk
#SingleInstance

updateMSStore() {
    ;; If Store is NOT open, open it and exit
    if !WinExist("Microsoft Store") {
        Run "ms-windows-store://updates"
        return
    }

    ;; If it IS open, activate and proceed
    WinActivate "Microsoft Store"
    Sleep 5000

    ;; OCR the window
    result := OCR.FromWindow("Microsoft Store")

    if !result {
        MsgBox "OCR failed"
        return 50
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

    ;; Re-scan for "Update all"
    result := OCR.FromWindow("Microsoft Store")

    if !result
        return

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

getStoreUpdateLogsToFile() {
    logDir := "C:\WinStoreUpdatesLog"
    logPath := logDir "\windowsStoreUpdatesLog.txt"

    ;; Ensure directory exists
    DirCreate logDir

    ;; Powershell command to save the log to directory. Store/Operational is where the MSApp store stores its logs. apparently. Change the maxevents flag for more or less logs.
    psCmd := "Get-WinEvent -LogName 'Microsoft-Windows-Store/Operational' -MaxEvents 500 | " .
        "Select-Object TimeCreated, Id, LevelDisplayName, Message | " .
        "Format-Table -Wrap | Out-String -Width 4096 | " .
        "Set-Content -Path '" logPath "' -Encoding UTF8"

    RunWait 'powershell -NoProfile -ExecutionPolicy Bypass -Command "' psCmd '"', , "Hide"

    ;; Error check if the file is written
    if !FileExist(logPath) {
        return 50
    }

    ;; Error check if the file is empty
    if (FileGetSize(logPath) = 0) {
        return 50
    }

    return 0
}

;; Initial launch
Run "ms-windows-store://updates"
Sleep 7000

;; Runs the command to update the store every 10 seconds
SetTimer(updateMSStore, 10000)

;; Runs this app for 15 minutes
Sleep 900000

;; Stores the updates
result := getStoreUpdateLogsToFile()

if (result = 50) {
    ExitApp 50
}

ExitApp 0