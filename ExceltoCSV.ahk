#Requires AutoHotkey v2.0
#SingleInstance Force

app := ExcelToCsvApp()

class ExcelToCsvApp {
    __New() {
        this.files := []
        this.folderPath := ""
        this.outDir := ""
        this.gui := Gui("+Resize -MaximizeBox", "Excel -> CSV Converter")
        this.BuildGui()
    }

    BuildGui() {
        g := this.gui
        g.SetFont("s10", "Segoe UI")

        g.Add("GroupBox", "x10 y10 w480 h150", "Input")
        this.rFiles := g.Add("Radio", "x25 y32 w440 h20 Checked", "Convert individual files")
        this.rFolder := g.Add("Radio", "x25 y56 w440 h20", "Convert a whole folder (recursive, includes subfolders)")
        this.rFiles.OnEvent("Click", (*) => this.UpdateMode())
        this.rFolder.OnEvent("Click", (*) => this.UpdateMode())

        this.lbFiles := g.Add("ListBox", "x25 y82 w350 h68 vFileListBox")
        this.btnAdd := g.Add("Button", "x385 y82 w90 h24", "Add Files...")
        this.btnRemove := g.Add("Button", "x385 y110 w90 h24", "Remove")
        this.btnClearFiles := g.Add("Button", "x385 y138 w90 h24", "Clear All")
        this.btnAdd.OnEvent("Click", (*) => this.AddFiles())
        this.btnRemove.OnEvent("Click", (*) => this.RemoveSelected())
        this.btnClearFiles.OnEvent("Click", (*) => this.ClearFiles())

        this.edFolder := g.Add("Edit", "x25 y82 w350 h24 ReadOnly Hidden")
        this.btnBrowseFolder := g.Add("Button", "x385 y82 w90 h24 Hidden", "Browse...")
        this.btnBrowseFolder.OnEvent("Click", (*) => this.BrowseFolder())

        g.Add("GroupBox", "x10 y168 w480 h56", "Output Folder")
        this.edOutput := g.Add("Edit", "x25 y190 w350 h24 ReadOnly")
        this.btnBrowseOutput := g.Add("Button", "x385 y190 w90 h24", "Browse...")
        this.btnBrowseOutput.OnEvent("Click", (*) => this.BrowseOutput())

        this.progress := g.Add("Progress", "x10 y234 w480 h18 Range0-100")
        this.txtStatus := g.Add("Text", "x10 y256 w480 h18", "Ready.")

        this.log := g.Add("Edit", "x10 y280 w480 h150 ReadOnly VScroll")

        this.btnConvert := g.Add("Button", "x10 y438 w150 h30 Default", "Convert")
        this.btnOpenOutput := g.Add("Button", "x170 y438 w150 h30 Disabled", "Open Output Folder")
        this.btnClose := g.Add("Button", "x340 y438 w150 h30", "Close")
        this.btnConvert.OnEvent("Click", (*) => this.StartConversion())
        this.btnOpenOutput.OnEvent("Click", (*) => this.OpenOutputFolder())
        this.btnClose.OnEvent("Click", (*) => ExitApp())

        g.OnEvent("Close", (*) => ExitApp())

        this.UpdateMode()
        g.Show("w500 h478")
    }

    UpdateMode() {
        filesMode := this.rFiles.Value
        for ctrl in [this.lbFiles, this.btnAdd, this.btnRemove, this.btnClearFiles]
            ctrl.Visible := filesMode
        for ctrl in [this.edFolder, this.btnBrowseFolder]
            ctrl.Visible := !filesMode
    }

    AddFiles() {
        picked := FileSelect("M3", , "Select Excel file(s) to convert", "Excel Files (*.xlsx; *.xlsm; *.xls)")
        if !IsObject(picked) || picked.Length = 0
            return
        for f in picked {
            isNew := true
            for existing in this.files {
                if existing = f {
                    isNew := false
                    break
                }
            }
            if isNew
                this.files.Push(f)
        }
        this.RefreshFileListBox()
    }

    RefreshFileListBox() {
        this.lbFiles.Delete()
        if this.files.Length
            this.lbFiles.Add(this.files)
    }

    RemoveSelected() {
        idx := this.lbFiles.Value
        if !idx
            return
        this.files.RemoveAt(idx)
        this.RefreshFileListBox()
    }

    ClearFiles() {
        this.files := []
        this.RefreshFileListBox()
    }

    BrowseFolder() {
        folder := DirSelect(, 3, "Select a folder containing Excel files")
        if folder = ""
            return
        this.folderPath := folder
        this.edFolder.Value := folder
    }

    BrowseOutput() {
        folder := DirSelect(, 3, "Select output folder for CSV files")
        if folder = ""
            return
        this.outDir := folder
        this.edOutput.Value := folder
    }

    OpenOutputFolder() {
        if this.outDir != ""
            Run(this.outDir)
    }

    AppendLog(line) {
        cur := this.log.Value
        this.log.Value := cur (cur = "" ? "" : "`r`n") line
        try {
            this.log.Focus()
            Send("^{End}")
        }
    }

    SetControlsEnabled(state) {
        this.rFiles.Enabled := state
        this.rFolder.Enabled := state
        this.btnAdd.Enabled := state && this.rFiles.Value
        this.btnRemove.Enabled := state && this.rFiles.Value
        this.btnClearFiles.Enabled := state && this.rFiles.Value
        this.btnBrowseFolder.Enabled := state && !this.rFiles.Value
        this.btnBrowseOutput.Enabled := state
        this.btnConvert.Enabled := state
    }

    StartConversion() {
        fileList := []
        if this.rFiles.Value {
            if this.files.Length = 0 {
                MsgBox("Please add at least one Excel file.", "Excel -> CSV Converter", "Icon!")
                return
            }
            fileList := this.files.Clone()
        } else {
            if this.folderPath = "" {
                MsgBox("Please choose a folder.", "Excel -> CSV Converter", "Icon!")
                return
            }
            Loop Files, this.folderPath "\*.*", "R" {
                ext := StrLower(A_LoopFileExt)
                if (ext = "xlsx" || ext = "xlsm" || ext = "xls")
                    fileList.Push(A_LoopFileFullPath)
            }
            if fileList.Length = 0 {
                MsgBox("No Excel files found in that folder.", "Excel -> CSV Converter", "Icon!")
                return
            }
        }

        if this.outDir = "" {
            MsgBox("Please choose an output folder.", "Excel -> CSV Converter", "Icon!")
            return
        }

        this.SetControlsEnabled(false)
        this.btnOpenOutput.Enabled := false
        this.log.Value := ""
        this.progress.Opt("Range0-" fileList.Length)
        this.progress.Value := 0

        xl := ""
        try {
            xl := ComObject("Excel.Application")
        } catch {
            this.AppendLog("ERROR: Could not start Excel. Make sure Microsoft Excel is installed on this machine.")
            this.txtStatus.Text := "Failed."
            this.SetControlsEnabled(true)
            return
        }

        filesOk := 0
        sheetsOk := 0
        errors := []

        try {
            xl.Visible := false
            xl.DisplayAlerts := false
            xl.ScreenUpdating := false

            for index, filePath in fileList {
                this.txtStatus.Text := "Converting (" index "/" fileList.Length "): " filePath
                this.AppendLog("Processing: " filePath)
                Sleep(1)
                try {
                    exported := ConvertWorkbook(xl, filePath, this.outDir, 62, 6)
                    sheetsOk += exported
                    filesOk += 1
                    this.AppendLog("  -> OK (" exported " sheet(s) exported)")
                } catch as err {
                    errors.Push(filePath " -> " err.Message)
                    this.AppendLog("  -> FAILED: " err.Message)
                }
                this.progress.Value := index
                Sleep(1)
            }
        } finally {
            if IsObject(xl)
                xl.Quit()
            xl := ""
        }

        this.txtStatus.Text := "Done."
        this.AppendLog("----")
        this.AppendLog("Files processed: " filesOk " / " fileList.Length)
        this.AppendLog("CSV sheets exported: " sheetsOk)
        if errors.Length
            this.AppendLog("Errors: " errors.Length)

        this.SetControlsEnabled(true)
        this.btnOpenOutput.Enabled := true
    }
}

ConvertWorkbook(xl, filePath, outDir, xlCSVUTF8, xlCSV) {
    SplitPath(filePath, , , , &nameNoExt)
    baseName := SanitizeFileName(nameNoExt)

    wb := xl.Workbooks.Open(filePath, 0, true)
    exported := 0
    try {
        for sheet in wb.Worksheets {
            sheetName := SanitizeFileName(sheet.Name)
            outPath := GetUniquePath(outDir "\" baseName "_" sheetName ".csv")

            sheet.Copy()
            tempWb := xl.ActiveWorkbook
            try {
                try {
                    tempWb.SaveAs(outPath, xlCSVUTF8)
                } catch {
                    tempWb.SaveAs(outPath, xlCSV)
                }
                exported += 1
            } finally {
                tempWb.Close(false)
            }
        }
    } finally {
        wb.Close(false)
    }
    return exported
}

SanitizeFileName(name) {
    name := RegExReplace(name, '[\\/:\*\?"<>\|]', "_")
    name := Trim(name, " .")
    if name = ""
        name := "Sheet"
    return name
}

GetUniquePath(path) {
    if !FileExist(path)
        return path
    SplitPath(path, , &dir, &ext, &nameNoExt)
    n := 2
    Loop {
        candidate := dir "\" nameNoExt " (" n ")." ext
        if !FileExist(candidate)
            return candidate
        n += 1
    }
}
