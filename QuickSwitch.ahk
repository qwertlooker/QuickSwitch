#Requires AutoHotkey v2.0
#SingleInstance Force

;=============================================================================
;   QuickSwitch — AutoOnly  (AHK v2)
;
;   文件对话框激活时，自动跳转到当前打开的文件管理器所在目录
;   支持：Windows Explorer / Total Commander / XYPlorer / Directory Opus
;=============================================================================

Z_SCAN_MAX := 6

;=============================================================================
;   主循环
;=============================================================================

loop {
    WinWaitActive "ahk_class #32770"

    dlgHwnd := WinExist("A")
    dlgKind := DetectDialogKind(dlgHwnd)

    if (dlgKind != "") {
        folder := FindManagerFolder(dlgHwnd, Z_SCAN_MAX)
        if (IsRealFolder(folder)) {
            NavigateDialog(dlgHwnd, folder)
        }
    }

    ; 短暂等待，避免在对话框关闭前重复触发
    Sleep 30
    WinWaitNotActive "ahk_id " dlgHwnd
}

;=============================================================================
;   检测对话框类型
;=============================================================================

DetectDialogKind(hwnd) {
    local ctrls, c, has

    try {
        ctrls := WinGetControls("ahk_id " hwnd)
    } catch {
        return ""
    }

    has := Map()
    for c in ctrls {
        has[c] := true
    }

    if (!has.Has("Edit1") or !has.Has("ToolbarWindow321")) {
        return ""
    }
    if (has.Has("DirectUIHWND1")) {
        return "MODERN"
    }
    return ""
}

;=============================================================================
;   Z-Order 范围扫描，找第一个已知文件管理器
;=============================================================================

FindManagerFolder(dlgHwnd, scanMax) {
    local wins, dlgPos, i, offset, targetHwnd, targetClass, folder

    wins   := WinGetList()
    dlgPos := 0

    loop wins.Length {
        if (wins[A_Index] = dlgHwnd) {
            dlgPos := A_Index
            break
        }
    }

    if (dlgPos = 0) {
        return ""
    }

    loop scanMax {
        offset := A_Index
        i      := dlgPos + offset
        if (i > wins.Length) {
            break
        }

        targetHwnd  := wins[i]
        targetClass := ""

        try {
            targetClass := WinGetClass("ahk_id " targetHwnd)
        } catch {
            continue
        }

        switch targetClass {
            case "CabinetWClass":    folder := FolderFrom_Explorer(targetHwnd)
            case "TTOTAL_CMD":       folder := FolderFrom_TC(targetHwnd)
            case "ThunderRT6FormDC": folder := FolderFrom_XYPlorer(targetHwnd)
            default:                 continue
        }

        if (IsRealFolder(folder)) {
            return folder
        }
    }

    return ""
}

;=============================================================================
;   Explorer
;=============================================================================

FolderFrom_Explorer(hwnd) {
    local result

    result := ""
    try {
        for w in ComObject("Shell.Application").Windows {
            try {
                if (w.hwnd = hwnd) {
                    result := w.Document.Folder.Self.Path
                    break
                }
            } catch {
                continue
            }
        }
    } catch {
        return ""
    }
    return result
}

;=============================================================================
;   Total Commander
;=============================================================================

FolderFrom_TC(hwnd) {
    local saved, result

    saved       := ClipboardAll()
    A_Clipboard := ""
    result      := ""

    try {
        SendMessage 1075, 2029, 0,, "ahk_id " hwnd
        ClipWait 1
        result := A_Clipboard
    } catch {
        result := ""
    }

    A_Clipboard := saved
    return result
}

;=============================================================================
;   XYPlorer
;=============================================================================

FolderFrom_XYPlorer(hwnd) {
    local saved, result

    saved       := ClipboardAll()
    A_Clipboard := ""
    result      := ""

    try {
        XYSend(hwnd, "::copytext get('path', a);")
        ClipWait 2
        result := A_Clipboard
    } catch {
        result := ""
    }

    A_Clipboard := saved
    return result
}

XYSend(hwnd, msg) {
    local size, buf, cd

    size := StrLen(msg)
    buf  := Buffer(size * 2, 0)
    StrPut(msg, buf, "UTF-16")

    cd := Buffer(A_PtrSize * 3, 0)
    NumPut "Ptr",  4194305,  cd, 0
    NumPut "UInt", size * 2, cd, A_PtrSize
    NumPut "Ptr",  buf.Ptr,  cd, A_PtrSize * 2

    SendMessage 0x4A, 0, cd.Ptr,, "ahk_id " hwnd
}

;=============================================================================
;   导航对话框到目标文件夹
;
;   原理：路径末尾加 \ 写入 Edit1 并发 Enter
;         Windows 对话框识别到末尾有 \ 时视为目录导航而非文件名
;=============================================================================

NavigateDialog(hwnd, folder) {
    local navPath, oldText

    navPath := RTrim(folder, "\") "\"

    try {
        WinActivate "ahk_id " hwnd

        ; 轮询等待对话框真正成为前台，最多等 200ms
        loop 20 {
            if (WinActive("ahk_id " hwnd)) {
                break
            }
            Sleep 10
        }

        oldText := ControlGetText("Edit1", "ahk_id " hwnd)

        ControlSetText navPath, "Edit1", "ahk_id " hwnd
        ControlFocus   "Edit1", "ahk_id " hwnd
        ControlSend    "{Enter}", "Edit1", "ahk_id " hwnd

        ; 轮询等待对话框完成导航（Edit1 内容被清空或变化），最多等 500ms
        loop 50 {
            Sleep 10
            if (ControlGetText("Edit1", "ahk_id " hwnd) != navPath) {
                break
            }
        }

        ; 恢复原文件名
        ControlSetText oldText, "Edit1", "ahk_id " hwnd
        ControlFocus   "Edit1", "ahk_id " hwnd

    } catch {
    }
}

;=============================================================================
;   验证路径是否为真实存在的目录
;=============================================================================

IsRealFolder(path) {
    if (path = "" or StrLen(path) >= 260) {
        return false
    }
    return InStr(FileExist(path), "D") ? true : false
}
