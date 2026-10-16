#NoEnv
#SingleInstance Force
SendMode Input
SetWorkingDir %A_ScriptDir%

; Attach to prompt_library.xlsx if it is already open in a running Excel
; instance (e.g. the user has it open to review data), instead of always
; spawning a second hidden Excel process. A second process opening the same
; path shows Excel's "locked for editing" / opens read-only, which then
; fails on Save with 0x800A03EC. Returns an object:
;   {XL: Excel.Application, wb: Workbook, owned: true/false}
; owned=false means we attached to the user's own instance - callers must
; NOT Close()/Quit() it, only Save() it.
GetPromptWorkbook(xlsxPath, readOnly := false) {
    SplitPath, xlsxPath, xlsxName
    try {
        existingXL := ComObjActive("Excel.Application")
    } catch e {
        existingXL := ""
    }
    try {
        if IsObject(existingXL) {
            for wbk in existingXL.Workbooks {
                if (wbk.Name = xlsxName) {
                    result := {}
                    result["XL"] := existingXL
                    result["wb"] := wbk
                    result["owned"] := false
                    return result
                }
            }
        }
    } catch e {
        ; Instance is shutting down or unreachable: fall through and open our own
    }

    ; Not already open anywhere - open our own hidden instance
    XL := ComObjCreate("Excel.Application")
    XL.Visible := False
    try {
        ; readOnly opens with no owner lock file, so the library stays editable
        ; in Excel while the picker reads it (used for loading prompts)
        wb := XL.Workbooks.Open(xlsxPath, 0, readOnly)
    } catch e {
        ; Don't leak the hidden Excel process we just created (e.g. the file
        ; is genuinely locked by someone else and Open() itself threw)
        XL.Quit()
        throw e
    }
    result := {}
    result["XL"] := XL
    result["wb"] := wb
    result["owned"] := true
    return result
}

; Close a handle returned by GetPromptWorkbook(). Always safe to call.
ReleasePromptWorkbook(handle) {
    if !IsObject(handle)
        return
    if (handle["owned"]) {
        handle["wb"].Close(False)
        handle["XL"].Quit()
    }
    ; else: it's the user's own visible workbook - leave it open as-is
}

; Inlined Jxon.ahk content so this script is self-contained
Jxon_Load(ByRef src, args*)
{
	static q := Chr(34)

	key := "", is_key := false
	stack := [ tree := [] ]
	is_arr := { (tree): 1 }
	next := q . "{[01234567890-tfn"
	pos := 0
	while ( (ch := SubStr(src, ++pos, 1)) != "" )
	{
		if InStr(" `t`n`r", ch)
			continue
		if !InStr(next, ch, true)
		{
			ln := ObjLength(StrSplit(SubStr(src, 1, pos), "`n"))
			col := pos - InStr(src, "`n",, -(StrLen(src)-pos+1))

			msg := Format("{}: line {} col {} (char {})"
			,   (next == "")      ? ["Extra data", ch := SubStr(src, pos)][1]
			  : (next == "'")     ? "Unterminated string starting at"
			  : (next == "\")     ? "Invalid \escape"
			  : (next == ":")     ? "Expecting ':' delimiter"
			  : (next == q)       ? "Expecting object key enclosed in double quotes"
			  : (next == q . "}") ? "Expecting object key enclosed in double quotes or object closing '}'"
			  : (next == ",}")    ? "Expecting ',' delimiter or object closing '}'"
			  : (next == ",]")    ? "Expecting ',' delimiter or array closing ']'"
			  : [ "Expecting JSON value(string, number, [true, false, null], object or array)"
			    , ch := SubStr(src, pos, (SubStr(src, pos)~="[\]\},\s]|$")-1) ][1]
			, ln, col, pos)

			throw Exception(msg, -1, ch)
		}

		is_array := is_arr[obj := stack[1]]

		if i := InStr("{[", ch)
		{
			val := (proto := args[i]) ? new proto : {}
			is_array? ObjPush(obj, val) : obj[key] := val
			ObjInsertAt(stack, 1, val)
			
			is_arr[val] := !(is_key := ch == "{")
			next := q . (is_key ? "}" : "{[]0123456789-tfn")
		}

		else if InStr("}]", ch)
		{
			ObjRemoveAt(stack, 1)
			next := stack[1]==tree ? "" : is_arr[stack[1]] ? ",]" : ",}"
		}

		else if InStr(",:", ch)
		{
			is_key := (!is_array && ch == ",")
			next := is_key ? q : q . "{[01234567890-tfn"
		}

		else ; string | number | true | false | null
		{
			if (ch == q) ; string
			{
				i := pos
				while i := InStr(src, q,, i+1)
				{
					val := StrReplace(SubStr(src, pos+1, i-pos-1), "\\", "\u005C")
					static end := A_AhkVersion<"2" ? 0 : -1
					if (SubStr(val, end) != "\\")
						break
				}
				if !i ? (pos--, next := "'") : 0
					continue

				pos := i ; update pos

				  val := StrReplace(val,    "\/",  "/")
				, val := StrReplace(val, "" . q,    q)
				, val := StrReplace(val,    "\b", "`b")
				, val := StrReplace(val,    "\f", "`f")
				, val := StrReplace(val,    "\n", "`n")
				, val := StrReplace(val,    "\r", "`r")
				, val := StrReplace(val,    "\t", "`t")

				i := 0
				while i := InStr(val, "\",, i+1)
				{
					if (SubStr(val, i+1, 1) != "u") ? (pos -= StrLen(SubStr(val, i)), next := "\\") : 0
						continue 2

					; \uXXXX - JSON unicode escape sequence
					xxxx := Abs("0x" . SubStr(val, i+2, 4))
					if (A_IsUnicode || xxxx < 0x100)
						val := SubStr(val, 1, i-1) . Chr(xxxx) . SubStr(val, i+6)
				}

				if is_key
				{
					key := val, next := ":"
					continue
				}
			}

			else ; number | true | false | null
			{
				val := SubStr(src, pos, i := RegExMatch(src, "[\]\},\s]|$",, pos)-pos)
			
			; For numerical values, numerify integers and keep floats as is.
			; I'm not yet sure if I should numerify floats in v2.0-a ...
				static number := "number", integer := "integer"
				if val is %number%
				{
					if val is %integer%
						val += 0
				}
			; in v1.1, true,false,A_PtrSize,A_IsUnicode,A_Index,A_EventInfo,
			; SOMETIMES return strings due to certain optimizations. Since it
			; is just 'SOMETIMES', numerify to be consistent w/ v2.0-a
				else if (val == "true" || val == "false")
					val := %value% + 0
			; AHK_H has built-in null, can't do 'val := %value%' where value == "null"
			; as it would raise an exception in AHK_H(overriding built-in var)
				else if (val == "null")
					val := ""
			; any other values are invalid, continue to trigger error
					else if (pos--, next := "#")
						continue
				
				pos += i-1
			}
			
			is_array? ObjPush(obj, val) : obj[key] := val
			next := obj==tree ? "" : is_array ? ",]" : ",}"
		}
	}

	return tree[1]
}

Jxon_Dump(obj, indent:="", lvl:=1)
{
	static q := Chr(34)

	if IsObject(obj)
	{
		static Type := Func("Type")
		if Type ? (Type.Call(obj) != "Object") : (ObjGetCapacity(obj) == "")
			throw Exception("Object type not supported.", -1, Format("<Object at 0x{:p}>", &obj))

		is_array := 0
		for k in obj
			is_array := k == A_Index
		until !is_array

		static integer := "integer"
		if indent is %integer%
		{
			if (indent < 0)
				throw Exception("Indent parameter must be a postive integer.", -1, indent)
			spaces := indent, indent := ""
			Loop % spaces
				indent .= " "
		}
		indt := ""
		Loop, % indent ? lvl : 0
			indt .= indent

		lvl += 1, out := "" ; Make #Warn happy
		for k, v in obj
		{
			if IsObject(k) || (k == "")
				throw Exception("Invalid object key.", -1, k ? Format("<Object at 0x{:p}>", &obj) : "<blank>")
			
			if !is_array
				out .= ( ObjGetCapacity([k], 1) ? Jxon_Dump(k) : q . k . q ) ;// key
				    .  ( indent ? ": " : ":") ; token + padding
			out .= Jxon_Dump(v, indent, lvl) ; value
			    .  ( indent ? ",`n" . indt : "," ) ; token + indent
		}

		if (out != "")
		{
			out := Trim(out, ",`n" . indent)
			if (indent != "")
				out := "`n" . indt . out . "`n" . SubStr(indt, StrLen(indent)+1)
		}
		
		return is_array ? "[" . out . "]" : "{" . out . "}"
	}

	; Number
	else if (ObjGetCapacity([obj], 1) == "")
		return obj

	; String (null -> not supported by AHK)
	if (obj != "")
	{
		  obj := StrReplace(obj,  "\\",    "\\\\")
		, obj := StrReplace(obj,  "/",    "\\/")
		, obj := StrReplace(obj,    q, "\\" . q)
		, obj := StrReplace(obj, "`b",    "\b")
		, obj := StrReplace(obj, "`f",    "\f")
		, obj := StrReplace(obj, "`n",    "\n")
		, obj := StrReplace(obj, "`r",    "\r")
		, obj := StrReplace(obj, "`t",    "\t")

		static needle := (A_AhkVersion<"2" ? "O)" : "") . "[^\x20-\x7e]"
		while RegExMatch(obj, needle, m)
			obj := StrReplace(obj, m[0], Format("\\u{:04X}", Ord(m[0])))
	}
	
	return q . obj . q
}

; Global variables
GuiOpen := false
CurrentPrompt := ""
Prompts := []
PromptTitles := []
FilteredPromptTitles := []
Categories := []
CurrentFilter := "All"
LastActiveWindow := ""

; Caching variables
CachedFileTime := ""
PromptsLoaded := false

; Favorites variables
Favorites := {}
FavoritesFile := A_ScriptDir . "\favorites.ini"

; Editor variables
EditorMode := ""  ; "new" or "edit"
EditingPromptIndex := 0

; Settings variables
SettingsFile := A_ScriptDir . "\settings.ini"
CurrentHotkey := ""          ; No default - user must set on first run
HotkeyDisplay := ""          ; Human-readable version
NewHotkey := ""              ; Temporary storage during recording
NewHotkeyDisplay := ""       ; Temporary display version
IsRecording := false         ; Recording state
FirstRun := false            ; First run flag
OnboardingMode := false      ; True while the first-run guide is showing
PickerLoading := false       ; True while prompts are being read from Excel
OpenAfterLoad := false       ; Hotkey pressed while loading: open when ready

; Usage stats, sorting, and startup options
UsageFile := A_ScriptDir . "\usage.ini"
UsageCount := {}             ; promptKey -> number of inserts/copies
UsageLast := {}              ; promptKey -> last used timestamp (YYYYMMDDHH24MISS)
MinutesPerUse := 3           ; Used for the estimated time-saved figure
PendingPromptKey := ""       ; Key of the prompt being inserted (for stats)
PickerLoadStart := 0         ; A_TickCount when the current load began
StartupLink := A_Startup . "\AHK Prompt Picker.lnk"
StarterLibError := ""        ; Last error from CreateStarterLibrary (for the message box)

; Variable substitution
PromptVariables := []        ; List of unique variable names found
PendingPromptContent := ""   ; Content waiting for variable substitution

; Load settings and register hotkey at startup
Gosub, LoadSettings

; Check if first run (no hotkey set)
if (CurrentHotkey = "") {
    FirstRun := true
    Gosub, ShowFirstRunSetup
} else {
    Hotkey, %CurrentHotkey%, TogglePromptPicker, On UseErrorLevel
    if (ErrorLevel) {
        ; Saved hotkey is invalid: fall back to first-run setup instead of crashing
        MsgBox, 48, Hotkey Problem, The saved hotkey (%HotkeyDisplay%) could not be registered.`n`nPlease choose a new one.
        CurrentHotkey := ""
        HotkeyDisplay := ""
        FirstRun := true
        Gosub, ShowFirstRunSetup
    } else {
        SetTimer, PrewarmPrompts, -1500
    }
}

; Setup tray menu
Menu, Tray, NoStandard  ; Remove default menu items
Menu, Tray, Add, Open AHK Prompt Picker, TrayOpenPicker
Menu, Tray, Add  ; Separator
Menu, Tray, Add, Open Prompt Library in Excel, TrayOpenLibrary
Menu, Tray, Add, Usage Stats, TrayStats
Menu, Tray, Add  ; Separator
Menu, Tray, Add, Settings, TraySettings
Menu, Tray, Add, Help, TrayHelp
Menu, Tray, Add, Privacy, TrayPrivacy
Menu, Tray, Add  ; Separator
Menu, Tray, Add, Reload Script, TrayReload
Menu, Tray, Add, Exit, TrayExit
Menu, Tray, Default, Open AHK Prompt Picker  ; Double-click action
if (HotkeyDisplay != "")
    Menu, Tray, Tip, AHK Prompt Picker - %HotkeyDisplay%
else
    Menu, Tray, Tip, AHK Prompt Picker - No hotkey set

; Jump over the label definition to prevent auto-execution
Goto, SkipToggleLabel

; Toggle prompt picker (dynamically registered hotkey)
TogglePromptPicker:
; Don't open picker during first-run setup or while recording hotkey
if (FirstRun or IsRecording)
    return
; Only treat the picker as open if its window really exists. A stale GuiOpen
; flag used to make the first hotkey press after closing do nothing.
if (GuiOpen) {
    Gui, 1:+LastFoundExist
    if (WinExist()) {
        Gui, 1:Destroy
        GuiOpen := false
        ToolTip
        return
    }
    GuiOpen := false
}
; Remember the window that was active when hotkey was pressed
WinGet, LastActiveWindow, ID, A
; Prompts are still loading (Excel can take several seconds on a cold start).
; Tell the user, and open the picker automatically as soon as loading finishes.
; Safety net: if a previous load never finished (error or dialog), reset the flag
; so the hotkey can never stay dead for the rest of the session.
if (PickerLoading and A_TickCount - PickerLoadStart > 90000)
    PickerLoading := false
if (PickerLoading) {
    OpenAfterLoad := true
    ToolTip, Loading your prompts... the picker will open in a moment.
    return
}
PickerLoading := true
PickerLoadStart := A_TickCount
if (!PromptsLoaded)
    ToolTip, Loading your prompts... the picker will open in a moment.
Gosub, LoadFavorites
Gosub, LoadPromptsIfNeeded
ToolTip
PickerLoading := false
OpenAfterLoad := false
; Prompts could not be loaded (message already shown): try again on the next press
if (!PromptsLoaded)
    return
Gosub, CreatePromptPicker
return

; Load prompts in the background shortly after startup so the first hotkey
; press opens the picker instantly instead of waiting on Excel.
PrewarmPrompts:
if (PickerLoading or GuiOpen or FirstRun)
    return
PickerLoading := true
PickerLoadStart := A_TickCount
Gosub, LoadFavorites
Gosub, LoadPromptsIfNeeded
PickerLoading := false
if (OpenAfterLoad) {
    ; The user pressed the hotkey while we were still loading
    OpenAfterLoad := false
    ToolTip
    if (PromptsLoaded)
        Gosub, CreatePromptPicker
}
return

SkipToggleLabel:

; ==================== TRAY MENU HANDLERS ====================

TrayOpenPicker:
Gosub, TogglePromptPicker
return

TraySettings:
Gosub, ShowSettings
return

TrayHelp:
Gosub, ShowUserGuide
return

TrayPrivacy:
Gosub, ShowPrivacy
return

TrayOpenLibrary:
Gosub, OpenLibrary
return

TrayStats:
Gosub, ShowStats
return

TrayReload:
Reload
return

TrayExit:
ExitApp
return

; ==================== END TRAY MENU HANDLERS ====================

; Load favorites from INI file
LoadFavorites:
Favorites := {}
if FileExist(FavoritesFile) {
    IniRead, favSection, %FavoritesFile%, Favorites
    if (favSection != "ERROR") {
        Loop, Parse, favSection, `n
        {
            ; Each line is "promptId=1"
            RegExMatch(A_LoopField, "^(.+)=1$", match)
            if (match1 != "")
                Favorites[match1] := true
        }
    }
}
return

; Save favorites to INI file
SaveFavorites:
; Clear existing favorites section and rewrite
FileDelete, %FavoritesFile%
for promptId, isFav in Favorites {
    if (isFav)
        IniWrite, 1, %FavoritesFile%, Favorites, %promptId%
}
return

; Load settings from INI file
LoadSettings:
if FileExist(SettingsFile) {
    IniRead, savedHotkey, %SettingsFile%, Hotkey, Key, ^!p
    IniRead, savedDisplay, %SettingsFile%, Hotkey, Display, Ctrl+Alt+P
    CurrentHotkey := savedHotkey
    ; Ensure the key portion of display is uppercase (e.g., "Ctrl+Alt+p" -> "Ctrl+Alt+P")
    StringGetPos, lastPlus, savedDisplay, +, R
    if (lastPlus >= 0) {
        keyPart := SubStr(savedDisplay, lastPlus + 2)
        StringUpper, keyPart, keyPart
        HotkeyDisplay := SubStr(savedDisplay, 1, lastPlus + 1) . keyPart
    } else {
        HotkeyDisplay := savedDisplay
    }
}
IniRead, MinutesPerUse, %SettingsFile%, Stats, MinutesPerUse, 3
if MinutesPerUse is not integer
    MinutesPerUse := 3
if (MinutesPerUse < 1)
    MinutesPerUse := 3
return

; Save stats and view options to INI file
SaveOptionSettings:
IniWrite, %MinutesPerUse%, %SettingsFile%, Stats, MinutesPerUse
return

; Save settings to INI file
SaveSettings:
IniWrite, %CurrentHotkey%, %SettingsFile%, Hotkey, Key
IniWrite, %HotkeyDisplay%, %SettingsFile%, Hotkey, Display
return

; Check if Excel file changed and reload if needed
LoadPromptsIfNeeded:
xlsxPath := A_ScriptDir . "\prompt_library.xlsx"
FileGetTime, currentFileTime, %xlsxPath%, M

; Only reload if file changed or never loaded
if (!PromptsLoaded or currentFileTime != CachedFileTime) {
    CachedFileTime := currentFileTime
    Gosub, LoadPrompts
    PromptsLoaded := (PromptTitles.Length() > 0) ? true : false
    if (!PromptsLoaded)
        CachedFileTime := ""
}
return

; Load prompts from file
LoadPrompts:
Prompts := []
PromptTitles := []
FilteredPromptTitles := []
Categories := ["All", "Favorites", "Most used", "Recently used", "Archived"]
CurrentFilter := "All"

xlsxPath := A_ScriptDir . "\prompt_library.xlsx"
if !FileExist(xlsxPath) {
    MsgBox, 36, Prompt Library Not Found, AHK Prompt Picker could not find prompt_library.xlsx in:`n%A_ScriptDir%`n`nCreate a starter library with a few example prompts now?`n`nChoose No to exit, then place your existing prompt_library.xlsx in that folder and start the app again.
    IfMsgBox, No
        ExitApp
    if !CreateStarterLibrary(xlsxPath) {
        MsgBox, 16, Error, The starter library could not be created.`n`nDetails: %StarterLibError%`n`nMicrosoft Excel is required. You can also place a prompt_library.xlsx file in:`n%A_ScriptDir%
        ExitApp
    }
    Sleep, 1500  ; let the Excel instance that wrote the file finish closing
    MsgBox, 64, Starter Library Created, A starter prompt library was created:`n%xlsxPath%`n`nAdd your own prompts with the New button in the app, or use Open in Excel to edit the library directly.
}

; Use Excel COM to read titles from column B and content from column C (starting row 2)
; Attaches to prompt_library.xlsx if the user already has it open, instead of
; always spawning a second hidden Excel process (see GetPromptWorkbook above).
xlHandle := ""
loadFailed := false
try {
    xlHandle := GetPromptWorkbook(xlsxPath, true)
    XL := xlHandle["XL"]
    wb := xlHandle["wb"]
    sheet := wb.Sheets(1)
    ; Determine last used row in column B using Excel xlUp (-4162)
    lastRow := sheet.Cells(sheet.Rows.Count, 2).End(-4162).Row
    if (lastRow >= 2) {
        Loop, % (lastRow - 1) {
            row := A_Index + 1
            id := sheet.Cells(row, 1).Value        ; Column A (ID)
            title := sheet.Cells(row, 2).Value     ; Column B (Title)
            content := sheet.Cells(row, 3).Value   ; Column C (Content)
            category := sheet.Cells(row, 4).Value  ; Column D (Category)
            version := sheet.Cells(row, 5).Value   ; Column E (Version)
            created := sheet.Cells(row, 6).Value   ; Column F (Created)
            createdBy := sheet.Cells(row, 7).Value ; Column G (CreatedBy)
            modified := sheet.Cells(row, 8).Value  ; Column H (Modified)
            modifiedBy := sheet.Cells(row, 9).Value ; Column I (ModifiedBy)
            status := sheet.Cells(row, 10).Value   ; Column J (Status)
            tags := sheet.Cells(row, 11).Value     ; Column K (Tags)
            description := sheet.Cells(row, 12).Value ; Column L (Description)

            ; Normalize COM variants to strings
            if IsObject(id)
                id := id.Value
            if IsObject(title)
                title := title.Value
            if IsObject(content)
                content := content.Value
            if IsObject(category)
                category := category.Value
            if IsObject(version)
                version := version.Value
            if IsObject(created)
                created := created.Value
            if IsObject(createdBy)
                createdBy := createdBy.Value
            if IsObject(modified)
                modified := modified.Value
            if IsObject(modifiedBy)
                modifiedBy := modifiedBy.Value
            if IsObject(status)
                status := status.Value
            if IsObject(tags)
                tags := tags.Value
            if IsObject(description)
                description := description.Value

            ; Skip empty title rows (allow blanks elsewhere)
            if (title = "" or title = 0)
                continue

            title := title ""         ; force to string
            content := content ""     ; force to string
            id := id ""               ; force to string
            category := category ""   ; force to string
            version := version ""     ; force to string
            created := created ""     ; force to string
            createdBy := createdBy "" ; force to string
            modified := modified ""   ; force to string
            modifiedBy := modifiedBy "" ; force to string
            status := status ""       ; force to string
            tags := tags ""           ; force to string
            description := description "" ; force to string

            ; Default category if empty
            if (category = "" or category = "0")
                category := "Uncategorized"

            ; Default version if empty
            if (version = "" or version = "0")
                version := "1.0"

            ; Default status if empty
            if (status = "" or status = "0")
                status := "Active"

            ; Add category to list if not already present
            categoryExists := false
            Loop % Categories.Length() {
                if (Categories[A_Index] = category) {
                    categoryExists := true
                    break
                }
            }
            if (!categoryExists)
                Categories.Push(category)

            prompt := {}  ; create an object compatible with existing code
            prompt["id"] := id
            prompt["task"] := title
            prompt["content"] := content
            prompt["category"] := category
            prompt["version"] := version
            prompt["created"] := created
            prompt["createdBy"] := createdBy
            prompt["modified"] := modified
            prompt["modifiedBy"] := modifiedBy
            prompt["status"] := status
            prompt["tags"] := tags
            prompt["description"] := description
            ; Create a display string that includes the ID from column A
            ; Add status indicator for non-Active prompts
            display := (id != "") ? id . " " . title : title
            if (status = "Draft")
                display .= " [Draft]"
            else if (status = "Archived")
                display .= " [Archived]"
            prompt["_display"] := display

            Prompts.Push(prompt)
            PromptTitles.Push(display)
        }
    }
} catch e {
    loadFailed := true
    MsgBox, 16, Error, % "Failed reading prompt_library.xlsx:`n" . e.Message
}
; Cleanup Excel objects (leaves the user's own Excel window untouched if we attached to it)
ReleasePromptWorkbook(xlHandle)
xlHandle := ""
wb := ""
sheet := ""
XL := ""

if (PromptTitles.Length() = 0) {
    ; Keep the app running so the hotkey retries once the library is fixed
    if (!loadFailed)
        MsgBox, 48, No Prompts, No valid prompts found in:`n%xlsxPath%`n`nAdd prompts to the file, then press the hotkey again.
}
return

; Build the GUI interface
CreatePromptPicker:
Gui, 1:Destroy  ; Ensure clean slate before creating controls
; A g-label thread inherits the default GUI of the window that launched it.
; When launched from another window (e.g. a first-run dialog), unprefixed
; Gui/GuiControl commands would target that window instead of GUI 1 and the
; preview and metadata would stay blank. Always pin the default to GUI 1.
Gui, 1:Default
GuiOpen := true
Gosub, LoadUsage

; Title and tagline
Gui, Font, s12 Bold, Segoe UI
Gui, Add, Text, x10 y8 w210 h22, AHK Prompt Picker
Gui, Font, s8 Normal, Segoe UI
Gui, Add, Text, x10 y30 w200 h16 cGray, Your prompts, most apps, instantly.

; Filter dropdown
Gui, Font, s9, Segoe UI
Gui, Add, Text, x10 y52 w60 h20, Filter by:
CategoryListString := ""
Loop % Categories.Length() {
    if (A_Index > 1)
        CategoryListString .= "|"
    CategoryListString .= Categories[A_Index]
}
Gui, Add, DropDownList, x70 y49 w140 vCategoryFilter gFilterChanged Choose1, %CategoryListString%

; Tag search (filters the list below by substring match against each prompt's Tags field)
Gui, Add, Text, x10 y76 w60 h20, Search:
Gui, Add, Edit, x70 y73 w140 h22 vTagSearchBox gTagSearchChanged hwndhSearchBox
; Grey hint text inside the empty search box (EM_SETCUEBANNER)
DllCall("SendMessage", "Ptr", hSearchBox, "UInt", 0x1501, "Ptr", 1, "WStr", "Title, tags, content")

Gui, Add, Text, x10 y102 w200 h20, Available Prompts:

; Initialize filtered list (start with all prompts, Archived hidden by default -
; matches the "All" behavior in ApplyFilters; use the Archived filter to see them)
FilteredPromptTitles := []
Loop % Prompts.Length() {
    if (Prompts[A_Index]["status"] != "Archived")
        FilteredPromptTitles.Push(Prompts[A_Index]["_display"])
}

; Create the listbox with prompt titles
PromptListString := ""
Loop % FilteredPromptTitles.Length() {
    if (A_Index > 1)
        PromptListString .= "|"
    PromptListString .= FilteredPromptTitles[A_Index]
}
Gui, Add, ListBox, x10 y122 w200 h258 vPromptList gPromptSelected, %PromptListString%

; New, Duplicate, and Edit buttons below the listbox
Gui, Add, Button, x10 y385 w63 h28 gNewPrompt, New
Gui, Add, Button, x78 y385 w63 h28 gDuplicatePrompt, Duplicate
Gui, Add, Button, x146 y385 w64 h28 gEditPrompt, Edit

; Metadata GroupBox (aligned with Filter row)
Gui, Add, GroupBox, x220 y37 w520 h148, Prompt Info
; Row 1: Version | Created | Modified (dates stacked above "By" values)
Gui, Add, Text, x235 y55 w55 h16, Version:
Gui, Add, Text, x292 y55 w100 h16 vMetaVersion, -
Gui, Add, Text, x400 y55 w55 h16, Created:
Gui, Add, Text, x457 y55 w100 h16 vMetaCreated, -
Gui, Add, Text, x565 y55 w58 h16, Modified:
Gui, Add, Text, x625 y55 w105 h16 vMetaModified, -
; Row 2: Status | Created By | Modified By
Gui, Add, Text, x235 y73 w55 h16, Status:
Gui, Add, Text, x292 y73 w100 h16 vMetaStatus, -
Gui, Add, Text, x400 y73 w55 h16, Author:
Gui, Add, Text, x457 y73 w100 h16 vMetaCreatedBy, -
Gui, Add, Text, x565 y73 w58 h16, Editor:
Gui, Add, Text, x625 y73 w105 h16 vMetaModifiedBy, -
; Row 3: Category (full width)
Gui, Add, Text, x235 y91 w55 h16, Category:
Gui, Add, Text, x292 y91 w438 h16 vMetaCategory, -
; Row 4: Tags (full width)
Gui, Add, Text, x235 y109 w55 h16, Tags:
Gui, Add, Text, x292 y109 w438 h16 vMetaTags, -
; Row 5: Description label
Gui, Add, Text, x235 y127 w65 h16, Description:
; Row 6: Description Edit (full width, scrollable)
Gui, Add, Edit, x235 y143 w495 h40 vMetaDescription ReadOnly +VScroll

; Preview section
Gui, Add, Text, x220 y197 w200 h20, Prompt Preview
Gui, Add, Edit, x220 y217 w520 h190 vPreview ReadOnly +VScroll

; Button layout (with Favorite button)
Gui, Add, Button, x220 y420 w90 h30 gInsertContent Default, Insert Prompt
Gui, Add, Button, x315 y420 w70 h30 gCopyContent, Copy
Gui, Add, Button, x390 y420 w80 h30 gToggleFavorite vFavoriteBtn, Favorite
Gui, Add, Button, x475 y420 w60 h30 gClosePromptPicker, Close
Gui, Add, Button, x540 y420 w60 h30 gShowSettings, Settings
Gui, Add, Button, x605 y420 w55 h30 gShowUserGuide, Help
Gui, Add, Button, x665 y420 w70 h30 gShowPrivacy, Privacy

; Usage stats and library shortcut (top right)
Gui, Font, s9, Segoe UI
Gui, Add, Button, x540 y6 w90 h26 gShowStats, Usage Stats
Gui, Add, Button, x635 y6 w105 h26 gOpenLibrary, Open in Excel

; Credit line at bottom
Gui, Font, s8, Segoe UI
Gui, Add, Text, x10 y455 w200 h15 cGray, Created by Wesley Erickson
Gui, Add, Text, x610 y455 w140 h15 cGray +Right, v2.1 | Press %HotkeyDisplay%

Gui, Show, w760 h480, AHK Prompt Picker

; Reset filter, select the first prompt, and start the hover-description timer
CurrentFilter := "All"
Gosub, ApplyFilters
hv_lastItem := 0
SetTimer, ListHoverCheck, 250

; Watch the library file while the picker is open so edits made in Excel appear
; without restarting the script (checked every 2 seconds, applied after a save settles)
FileGetTime, LibWatchTime, %A_ScriptDir%\prompt_library.xlsx, M
LibWatchPending := ""
SetTimer, LibraryWatch, 2000
return

; Reload the list when prompt_library.xlsx was saved while the picker is open
LibraryWatch:
if (!GuiOpen) {
    SetTimer, LibraryWatch, Off
    return
}
; Do not disturb the editor, the variable dialog, or a load already in progress
Gui, 2:+LastFoundExist
if (WinExist() or PickerLoading)
    return
Gui, 4:+LastFoundExist
if (WinExist())
    return
FileGetTime, lw_now, %A_ScriptDir%\prompt_library.xlsx, M
if (lw_now = "" or lw_now = LibWatchTime) {
    LibWatchPending := ""
    return
}
; Wait until the timestamp is stable for one more check so a save in progress is not read half-written
if (lw_now != LibWatchPending) {
    LibWatchPending := lw_now
    return
}
LibWatchTime := lw_now
LibWatchPending := ""
PickerLoading := true
PickerLoadStart := A_TickCount
CachedFileTime := ""
PromptsLoaded := false
Gosub, LoadPromptsIfNeeded
PickerLoading := false
if (PromptsLoaded and GuiOpen) {
    Gosub, RefreshPromptList
    ToolTip, Prompt library updated
    SetTimer, RemoveTooltip, -1500
}
return

; Hover over a prompt in the list to see its description in a tooltip
ListHoverCheck:
if (!GuiOpen) {
    SetTimer, ListHoverCheck, Off
    if (hv_lastItem)
        ToolTip,,,, 2
    hv_lastItem := 0
    return
}
MouseGetPos, , , hv_win, hv_ctrl, 2
GuiControlGet, hv_lb, 1:Hwnd, PromptList
hv_item := 0
if (hv_ctrl = hv_lb and hv_lb != "") {
    VarSetCapacity(hv_pt, 8, 0)
    DllCall("GetCursorPos", "Ptr", &hv_pt)
    DllCall("ScreenToClient", "Ptr", hv_lb, "Ptr", &hv_pt)
    hv_x := NumGet(hv_pt, 0, "Int")
    hv_y := NumGet(hv_pt, 4, "Int")
    hv_res := DllCall("SendMessage", "Ptr", hv_lb, "UInt", 0x1A9, "Ptr", 0, "Ptr", (hv_y << 16) | (hv_x & 0xFFFF), "Ptr")  ; LB_ITEMFROMPOINT
    hv_idx := hv_res & 0xFFFF
    hv_top := DllCall("SendMessage", "Ptr", hv_lb, "UInt", 0x18E, "Ptr", 0, "Ptr", 0, "Ptr")  ; LB_GETTOPINDEX
    hv_h := DllCall("SendMessage", "Ptr", hv_lb, "UInt", 0x1A1, "Ptr", 0, "Ptr", 0, "Ptr")   ; LB_GETITEMHEIGHT
    ; Ignore the empty area below the last item
    if (hv_idx < FilteredPromptTitles.Length() and hv_h > 0 and hv_y < (hv_idx - hv_top + 1) * hv_h and hv_y >= 0)
        hv_item := hv_idx + 1
}
if (hv_item = hv_lastItem)
    return
hv_lastItem := hv_item
if (hv_item = 0) {
    ToolTip,,,, 2
    return
}
hv_display := FilteredPromptTitles[hv_item]
hv_desc := ""
Loop % Prompts.Length() {
    if (Prompts[A_Index]["_display"] = hv_display) {
        hv_desc := Prompts[A_Index]["description"]
        break
    }
}
if (hv_desc = "")
    hv_desc := "No description for this prompt."
ToolTip, % WrapTooltip(hv_display . "`n" . hv_desc, 60),,, 2
return

; When a new prompt is selected from the list
PromptSelected:
Gui, Submit, NoHide
CurrentPrompt := PromptList
Gosub, UpdatePreview
return

; When the category filter changes
FilterChanged:
Gui, Submit, NoHide
CurrentFilter := CategoryFilter
Gosub, ApplyFilters
return

; When the tag search box changes (fires live as the user types)
TagSearchChanged:
Gosub, ApplyFilters
return

; Open prompt_library.xlsx in Excel (the picker attaches to the open copy when it reloads)
OpenLibrary:
u_xlsx := A_ScriptDir . "\prompt_library.xlsx"
if !FileExist(u_xlsx) {
    MsgBox, 48, Prompt Library, prompt_library.xlsx was not found in:`n%A_ScriptDir%
    return
}

; 1. Already open in a running Excel (even a hidden one)? Show that copy.
u_done := false
try {
    u_xl := ComObjActive("Excel.Application")
    for u_wb in u_xl.Workbooks {
        if (u_wb.Name = "prompt_library.xlsx") {
            u_xl.Visible := true
            u_wb.Activate()
            u_hwnd := u_xl.Hwnd
            WinActivate, ahk_id %u_hwnd%
            u_done := true
            break
        }
    }
}
u_xl := ""
u_wb := ""
if (u_done)
    return

; 2. The picker may still be releasing its own background read of the file.
;    Wait up to 5 seconds for the file lock to clear before opening Excel.
ToolTip, Opening prompt library...
u_free := false
Loop, 10 {
    u_fh := FileOpen(u_xlsx, "a")
    if IsObject(u_fh) {
        u_fh.Close()
        u_free := true
        break
    }
    Sleep, 500
}
ToolTip
if (!u_free) {
    MsgBox, 48, Prompt Library, prompt_library.xlsx is locked by another Excel window or user.`n`nClose it there and try again. If nothing else is open, end any leftover Excel process in Task Manager.
    return
}
Run, "%u_xlsx%",, UseErrorLevel
if (ErrorLevel)
    MsgBox, 48, Prompt Library, The file could not be opened. Check that Microsoft Excel is installed.
return

; ==================== USAGE STATS ====================

; Load usage counts and last-used times from usage.ini
LoadUsage:
UsageCount := {}
UsageLast := {}
if FileExist(UsageFile) {
    IniRead, u_uses, %UsageFile%, Uses
    if (u_uses != "ERROR") {
        Loop, Parse, u_uses, `n, `r
        {
            u_pos := InStr(A_LoopField, "=")
            if (u_pos > 1)
                UsageCount[SubStr(A_LoopField, 1, u_pos - 1)] := SubStr(A_LoopField, u_pos + 1) + 0
        }
    }
    IniRead, u_last, %UsageFile%, LastUsed
    if (u_last != "ERROR") {
        Loop, Parse, u_last, `n, `r
        {
            u_pos := InStr(A_LoopField, "=")
            if (u_pos > 1)
                UsageLast[SubStr(A_LoopField, 1, u_pos - 1)] := SubStr(A_LoopField, u_pos + 1)
        }
    }
}
return

; Show the usage stats window
ShowStats:
Gosub, LoadUsage
Gui, Stats:Destroy

u_total := 0
u_used := 0
for u_k, u_n in UsageCount {
    u_total += u_n
    u_used += 1
}
u_count := Prompts.Length()
u_saved := FormatMinutes(u_total * MinutesPerUse)
IniRead, u_since, %UsageFile%, Meta, Since, NONE
if (u_since = "NONE") {
    u_sinceText := "no usage recorded yet"
} else {
    FormatTime, u_sinceText, %u_since%, MMM d, yyyy
}

Gui, Stats:Font, s11 Bold, Segoe UI
Gui, Stats:Add, Text, x15 y10 w490 h24, Your Prompt Picker Usage
Gui, Stats:Font, s9 Normal, Segoe UI
Gui, Stats:Add, Text, x15 y36 w490 h18 cGray, Tracking since: %u_sinceText%
Gui, Stats:Add, GroupBox, x15 y60 w490 h84, Summary
Gui, Stats:Add, Text, x30 y82 w460 h58, Total uses (inserts and copies): %u_total%`nPrompts used: %u_used% of %u_count%`nEstimated time saved: %u_saved% (at %MinutesPerUse% min per use, change in Settings)

if (u_total = 0) {
    Gui, Stats:Add, Text, x15 y160 w490 h40, No usage yet. Insert or copy a prompt and it will show up here.
} else {
    Gui, Stats:Add, ListView, x15 y156 w490 h250 Grid vStatsList, Prompt|Uses|Last Used
    Gui, Stats:Default
    u_titles := {}
    Loop % Prompts.Length() {
        u_p := Prompts[A_Index]
        u_key := UsageKey(u_p)
        u_titles[u_key] := u_p["_display"]
    }
    for u_k, u_n in UsageCount {
        u_name := u_titles.HasKey(u_k) ? u_titles[u_k] : u_k . " (removed)"
        u_lastText := ""
        if (UsageLast.HasKey(u_k) and UsageLast[u_k] != "")
            FormatTime, u_lastText, % UsageLast[u_k], MM/dd/yyyy h:mm tt
        LV_Add("", u_name, u_n, u_lastText)
    }
    LV_ModifyCol(1, 280)
    LV_ModifyCol(2, "50 Integer")
    LV_ModifyCol(3, 140)
    LV_ModifyCol(2, "SortDesc")
}

Gui, Stats:Font, s8, Segoe UI
Gui, Stats:Add, Text, x15 y412 w490 h30 cGray, Stored only on this computer in usage.ini. Only prompt IDs, counts, and times are recorded, never prompt text or the app you pasted into.
Gui, Stats:Font, s9, Segoe UI
Gui, Stats:Add, Button, x15 y448 w100 h30 gStatsReset, Reset Stats
Gui, Stats:Add, Button, x405 y448 w100 h30 gStatsClose Default, Close
Gui, Stats:Show, w520 h492, Usage Stats
return

StatsReset:
MsgBox, 292, Reset Stats, This clears all usage counts and times.`n`nContinue?
IfMsgBox, Yes
{
    FileDelete, %UsageFile%
    UsageCount := {}
    UsageLast := {}
    Gosub, ShowStats
}
return

StatsClose:
StatsGuiClose:
StatsGuiEscape:
Gui, Stats:Destroy
return

; ==================== END USAGE STATS ====================

; Rebuild FilteredPromptTitles from CurrentFilter (category/favorites) AND the
; tag search box (case-insensitive substring match against each prompt's Tags
; field), then refresh the ListBox. Shared by FilterChanged and TagSearchChanged
; so the two filters always combine instead of one silently overriding the other.
ApplyFilters:
GuiControlGet, TagSearchText, 1:, TagSearchBox
TagSearchText := Trim(TagSearchText)

FilteredPromptTitles := []
filterIdx := []
Loop % Prompts.Length() {
    prompt := Prompts[A_Index]

    isArchived := (prompt["status"] = "Archived")

    if (CurrentFilter = "Archived")
        categoryMatch := isArchived
    else if (CurrentFilter = "All")
        categoryMatch := !isArchived
    else if (CurrentFilter = "Favorites")
        categoryMatch := (Favorites[prompt["id"]] and !isArchived) ? true : false
    else if (CurrentFilter = "Most used") {
        usageKey := UsageKey(prompt)
        categoryMatch := (UsageCount[usageKey] > 0 and !isArchived) ? true : false
    } else if (CurrentFilter = "Recently used") {
        usageKey := UsageKey(prompt)
        categoryMatch := (UsageLast[usageKey] != "" and !isArchived) ? true : false
    } else
        categoryMatch := (prompt["category"] = CurrentFilter and !isArchived)

    if !categoryMatch
        continue

    if (TagSearchText != "") and !PromptMatchesSearch(prompt, TagSearchText)
        continue

    filterIdx.Push(A_Index)
}

; Most used / Recently used filters also set the order; other filters keep sheet order
filterIdx := SortPromptIndexes(filterIdx)
Loop % filterIdx.Length()
    FilteredPromptTitles.Push(Prompts[filterIdx[A_Index]]["_display"])

; Rebuild the listbox string
PromptListString := ""
Loop % FilteredPromptTitles.Length() {
    if (A_Index > 1)
        PromptListString .= "|"
    PromptListString .= FilteredPromptTitles[A_Index]
}

; Update the ListBox control
GuiControl, 1:, PromptList, |%PromptListString%

; Select first item if list is not empty
if (FilteredPromptTitles.Length() > 0) {
    CurrentPrompt := FilteredPromptTitles[1]
    GuiControl, 1:Choose, PromptList, 1
    Gosub, UpdatePreview
} else {
    CurrentPrompt := ""
    GuiControl, 1:, MetaVersion, -
    GuiControl, 1:, MetaStatus, -
    GuiControl, 1:, MetaTags, -
    GuiControl, 1:, MetaCategory, -
    GuiControl, 1:, MetaCreated, -
    GuiControl, 1:, MetaCreatedBy, -
    GuiControl, 1:, MetaModified, -
    GuiControl, 1:, MetaModifiedBy, -
    GuiControl, 1:, MetaDescription, -
    if (CurrentFilter = "Most used" or CurrentFilter = "Recently used")
        GuiControl, 1:, Preview, No usage yet. Insert or copy a prompt and it will show up here.
    else
        GuiControl, 1:, Preview, No prompts match this filter.
}
return

; Display content in preview
UpdatePreview:
PreviewText := ""
currentPromptId := ""
metaVersion := "-"
metaStatus := "-"
metaTags := "-"
metaCategory := "-"
metaCreated := "-"
metaCreatedBy := "-"
metaModified := "-"
metaModifiedBy := "-"
metaDescription := "-"

Loop % Prompts.Length() {
    if (Prompts[A_Index]["_display"] = CurrentPrompt) {
        prompt := Prompts[A_Index]
        currentPromptId := prompt["id"]

        ; Update metadata values
        metaVersion := prompt["version"]
        if RegExMatch(metaVersion, "^\d+\.\d+$") {
            metaVersion := RegExReplace(metaVersion, "(\.\d*?)0+$", "$1")
            metaVersion := RegExReplace(metaVersion, "\.$", ".0")
        } else if RegExMatch(metaVersion, "^\d+$")
            metaVersion := metaVersion ".0"
        if (metaVersion = "")
            metaVersion := "-"
        metaStatus := (prompt["status"] != "") ? prompt["status"] : "-"
        metaTags := (prompt["tags"] != "") ? prompt["tags"] : "-"
        metaCategory := (prompt["category"] != "") ? prompt["category"] : "-"
        metaDescription := (prompt["description"] != "") ? prompt["description"] : "-"

        ; Format created date as MM/DD/YYYY (handle YYYY-MM-DD or M/D/YYYY, strip time)
        if (prompt["created"] != "") {
            createdRaw := prompt["created"]
            if (RegExMatch(createdRaw, "(\d{4})-(\d{2})-(\d{2})", m))
                metaCreated := m2 "/" m3 "/" m1
            else if (RegExMatch(createdRaw, "(\d{1,2})/(\d{1,2})/(\d{4})", m))
                metaCreated := Format("{:02}/{:02}/{}", m1, m2, m3)
            else
                metaCreated := createdRaw
        }
        if (prompt["createdBy"] != "")
            metaCreatedBy := prompt["createdBy"]

        ; Format modified date as MM/DD/YYYY (handle YYYY-MM-DD or M/D/YYYY, strip time)
        if (prompt["modified"] != "") {
            modifiedRaw := prompt["modified"]
            if (RegExMatch(modifiedRaw, "(\d{4})-(\d{2})-(\d{2})", m))
                metaModified := m2 "/" m3 "/" m1
            else if (RegExMatch(modifiedRaw, "(\d{1,2})/(\d{1,2})/(\d{4})", m))
                metaModified := Format("{:02}/{:02}/{}", m1, m2, m3)
            else
                metaModified := modifiedRaw
        }
        if (prompt["modifiedBy"] != "")
            metaModifiedBy := prompt["modifiedBy"]

        ; Show only the prompt content in preview
        PreviewText := prompt["content"]
        break
    }
}

; Update metadata labels
GuiControl, 1:, MetaVersion, %metaVersion%
GuiControl, 1:, MetaStatus, %metaStatus%
GuiControl, 1:, MetaTags, %metaTags%
GuiControl, 1:, MetaCategory, %metaCategory%
GuiControl, 1:, MetaCreated, %metaCreated%
GuiControl, 1:, MetaCreatedBy, %metaCreatedBy%
GuiControl, 1:, MetaModified, %metaModified%
GuiControl, 1:, MetaModifiedBy, %metaModifiedBy%
GuiControl, 1:, MetaDescription, %metaDescription%
GuiControl, 1:, Preview, %PreviewText%

; Update favorite button text
if (currentPromptId != "" and Favorites[currentPromptId])
    GuiControl, 1:, FavoriteBtn, Unfavorite
else
    GuiControl, 1:, FavoriteBtn, Favorite
return

; Toggle favorite status for current prompt
ToggleFavorite:
if (CurrentPrompt = "")
    return

; Find the current prompt's ID
currentPromptId := ""
Loop % Prompts.Length() {
    if (Prompts[A_Index]["_display"] = CurrentPrompt) {
        currentPromptId := Prompts[A_Index]["id"]
        break
    }
}

if (currentPromptId = "")
    return

; Toggle favorite status
if (Favorites[currentPromptId]) {
    Favorites.Delete(currentPromptId)
    GuiControl,, FavoriteBtn, Add Favorite
    ToolTip, Removed from favorites
} else {
    Favorites[currentPromptId] := true
    GuiControl,, FavoriteBtn, Unfavorite
    ToolTip, Added to favorites!
}

; Save to file
Gosub, SaveFavorites
SetTimer, RemoveTooltip, -1000

; If currently viewing Favorites filter, refresh the list
Gui, Submit, NoHide
if (CategoryFilter = "Favorites")
    Gosub, FilterChanged
return

; Open editor for new prompt
NewPrompt:
EditorMode := "new"
EditingPromptIndex := 0

; Generate next ID based on highest existing numeric ID
nextId := "001"
maxId := 0
Loop % Prompts.Length() {
    promptId := Prompts[A_Index]["id"]
    ; Extract numeric portion from ID (handles "001", "39", "MTG-042", etc.)
    if RegExMatch(promptId, "(\d+)", idMatch) {
        numId := idMatch1 + 0  ; Convert to number
        if (numId > maxId)
            maxId := numId
    }
}
nextId := Format("{:03}", maxId + 1)

Gosub, ShowPromptEditor
GuiControl, 2:, EditorId, %nextId%
GuiControl, 2:, EditorTitle,
GuiControl, 2:, EditorContent,
GuiControl, 2:, EditorCategory,
GuiControl, 2:, EditorTags,
GuiControl, 2:, EditorDescription,
GuiControl, 2:Choose, EditorStatus, 1  ; Default to "Active"
GuiControl, 2:, EditorVersionDisplay, 1.0
GuiControl, 2:, EditorCreatedDisplay, (new)
GuiControl, 2:, EditorModifiedDisplay, -
return

; Duplicate selected prompt as a new one
DuplicatePrompt:
if (CurrentPrompt = "")
    return

EditorMode := "new"
EditingPromptIndex := 0

; Generate next ID based on highest existing numeric ID
nextId := "001"
maxId := 0
Loop % Prompts.Length() {
    promptId := Prompts[A_Index]["id"]
    ; Extract numeric portion from ID (handles "001", "39", "MTG-042", etc.)
    if RegExMatch(promptId, "(\d+)", idMatch) {
        numId := idMatch1 + 0  ; Convert to number
        if (numId > maxId)
            maxId := numId
    }
}
nextId := Format("{:03}", maxId + 1)

; Find the current prompt to copy from
Loop % Prompts.Length() {
    if (Prompts[A_Index]["_display"] = CurrentPrompt) {
        prompt := Prompts[A_Index]

        Gosub, ShowPromptEditor
        GuiControl, 2:, EditorId, %nextId%
        GuiControl, 2:, EditorTitle, % "Copy of " . prompt["task"]
        GuiControl, 2:, EditorContent, % prompt["content"]
        GuiControl, 2:, EditorCategory, % prompt["category"]
        GuiControl, 2:, EditorTags, % prompt["tags"]
        GuiControl, 2:, EditorDescription, % prompt["description"]
        GuiControl, 2:Choose, EditorStatus, 1  ; Default to Active
        GuiControl, 2:, EditorVersionDisplay, 1.0
        GuiControl, 2:, EditorCreatedDisplay, (new)
        GuiControl, 2:, EditorModifiedDisplay, -
        break
    }
}
return

; Open editor for existing prompt
EditPrompt:
if (CurrentPrompt = "")
    return

EditorMode := "edit"

; Find the current prompt
Loop % Prompts.Length() {
    if (Prompts[A_Index]["_display"] = CurrentPrompt) {
        EditingPromptIndex := A_Index
        prompt := Prompts[A_Index]

        Gosub, ShowPromptEditor
        GuiControl, 2:, EditorId, % prompt["id"]
        GuiControl, 2:, EditorTitle, % prompt["task"]
        GuiControl, 2:, EditorContent, % prompt["content"]
        GuiControl, 2:, EditorCategory, % prompt["category"]
        GuiControl, 2:, EditorTags, % prompt["tags"]
        GuiControl, 2:, EditorDescription, % prompt["description"]

        ; Set status dropdown
        promptStatus := prompt["status"]
        if (promptStatus = "Draft")
            GuiControl, 2:Choose, EditorStatus, 2
        else if (promptStatus = "Archived")
            GuiControl, 2:Choose, EditorStatus, 3
        else
            GuiControl, 2:Choose, EditorStatus, 1  ; Active (default)

        ; Display metadata (dates formatted as MM/DD/YYYY)
        verInfo := prompt["version"]
        if RegExMatch(verInfo, "^\d+\.\d+$") {
            verInfo := RegExReplace(verInfo, "(\.\d*?)0+$", "$1")
            verInfo := RegExReplace(verInfo, "\.$", ".0")
        } else if RegExMatch(verInfo, "^\d+$")
            verInfo := verInfo ".0"
        if (verInfo = "")
            verInfo := "1.0"
        GuiControl, 2:, EditorVersionDisplay, %verInfo%

        ; Format created date (handle YYYY-MM-DD or M/D/YYYY formats, strip time)
        createdInfo := "-"
        if (prompt["created"] != "") {
            createdRaw := prompt["created"]
            if (RegExMatch(createdRaw, "(\d{4})-(\d{2})-(\d{2})", m))
                createdInfo := m2 "/" m3 "/" m1
            else if (RegExMatch(createdRaw, "(\d{1,2})/(\d{1,2})/(\d{4})", m))
                createdInfo := Format("{:02}/{:02}/{}", m1, m2, m3)
            else
                createdInfo := createdRaw
        }
        GuiControl, 2:, EditorCreatedDisplay, %createdInfo%

        ; Format modified date (handle YYYY-MM-DD or M/D/YYYY formats, strip time)
        modifiedInfo := "-"
        if (prompt["modified"] != "") {
            modifiedRaw := prompt["modified"]
            if (RegExMatch(modifiedRaw, "(\d{4})-(\d{2})-(\d{2})", m))
                modifiedInfo := m2 "/" m3 "/" m1
            else if (RegExMatch(modifiedRaw, "(\d{1,2})/(\d{1,2})/(\d{4})", m))
                modifiedInfo := Format("{:02}/{:02}/{}", m1, m2, m3)
            else
                modifiedInfo := modifiedRaw
        }
        GuiControl, 2:, EditorModifiedDisplay, %modifiedInfo%
        break
    }
}
return

; Show the prompt editor dialog
ShowPromptEditor:
Gui, 2:Destroy
Gui, 2:Font, s9, Segoe UI

; Row 1: ID and Category
Gui, 2:Add, Text, x10 y10 w80 h20, ID:
Gui, 2:Add, Edit, x85 y7 w70 h24 vEditorId
Gui, 2:Add, Text, x163 y10 w60 h20, Category:
Gui, 2:Add, Edit, x225 y7 w140 h24 vEditorCategory
Gui, 2:Add, Text, x375 y10 w50 h20, Status:
Gui, 2:Add, DropDownList, x425 y7 w100 r3 vEditorStatus Choose1, Active|Draft|Archived

; Row 2: Title
Gui, 2:Add, Text, x10 y45 w80 h20, Title:
Gui, 2:Add, Edit, x85 y42 w440 h24 vEditorTitle

; Row 3: Tags
Gui, 2:Add, Text, x10 y80 w80 h20, Tags:
Gui, 2:Add, Edit, x85 y77 w440 h24 vEditorTags

; Row 4: Description (multi-line with scroll)
Gui, 2:Add, Text, x10 y112 w80 h20, Description:
Gui, 2:Add, Edit, x85 y109 w440 h50 vEditorDescription +Multi +VScroll

; Row 5: Content with formatting toolbar
Gui, 2:Add, Text, x10 y169 w55 h20, Content:
Gui, 2:Font, s8, Segoe UI
Gui, 2:Add, Button, x70 y166 w28 h22 gFmtBold, B
Gui, 2:Add, Button, x100 y166 w28 h22 gFmtItalic, I
Gui, 2:Add, Button, x130 y166 w28 h22 gFmtCode, ``
Gui, 2:Add, Button, x160 y166 w38 h22 gFmtLink, Link
Gui, 2:Add, Button, x200 y166 w28 h22 gFmtH1, H1
Gui, 2:Add, Button, x230 y166 w28 h22 gFmtH2, H2
Gui, 2:Add, Button, x260 y166 w28 h22 gFmtBullet, -
Gui, 2:Add, Button, x290 y166 w28 h22 gFmtQuote, >
Gui, 2:Add, Button, x320 y166 w32 h22 gFmtCodeBlock, ```
Gui, 2:Add, Button, x354 y166 w38 h22 gFmtVar, {{}}
Gui, 2:Font, s9, Segoe UI
Gui, 2:Add, Edit, x10 y189 w515 h130 vEditorContent +Multi +WantReturn +VScroll

; Variable tip box
Gui, 2:Font, s8, Segoe UI
Gui, 2:Add, GroupBox, x10 y321 w515 h40, Tip: Dynamic Variables
Gui, 2:Add, Text, x20 y336 w495 h20 cGray, Use {{variable_name}} to create fill-in fields. Example: {{customer_name}}, {{date}}

; Row 6: Metadata (read-only display) - simplified with date only
Gui, 2:Add, Text, x10 y369 w55 h18, Version:
Gui, 2:Add, Text, x67 y369 w100 h18 vEditorVersionDisplay, 1.0
Gui, 2:Add, Text, x180 y369 w55 h18, Created:
Gui, 2:Add, Text, x237 y369 w105 h18 vEditorCreatedDisplay, -
Gui, 2:Add, Text, x355 y369 w58 h18, Modified:
Gui, 2:Add, Text, x415 y369 w110 h18 vEditorModifiedDisplay, -
Gui, 2:Font, s9, Segoe UI

; Row 7: Buttons
Gui, 2:Add, Button, x10 y397 w100 h30 gSavePrompt Default, Save
Gui, 2:Add, Button, x120 y397 w100 h30 gCancelEditor, Cancel
if (EditorMode = "edit") {
    Gui, 2:Add, Button, x425 y397 w100 h30 gDeletePrompt, Delete
}

editorTitle := (EditorMode = "new") ? "New Prompt" : "Edit Prompt"
Gui, 2:Show, w535 h440, %editorTitle%
return

; Cancel editor
CancelEditor:
2GuiClose:
2GuiEscape:
Gui, 2:Destroy
return

; Save prompt to Excel
SavePrompt:
Gui, 2:Submit, NoHide
GuiControlGet, newId, 2:, EditorId
GuiControlGet, newTitle, 2:, EditorTitle
GuiControlGet, newContent, 2:, EditorContent
GuiControlGet, newCategory, 2:, EditorCategory
GuiControlGet, newStatus, 2:, EditorStatus
GuiControlGet, newTags, 2:, EditorTags
GuiControlGet, newDescription, 2:, EditorDescription

; Validate required fields
if (newTitle = "") {
    MsgBox, 48, Validation Error, Title is required.
    return
}
if (newContent = "") {
    MsgBox, 48, Validation Error, Content is required.
    return
}

; Default category if empty
if (newCategory = "")
    newCategory := "Uncategorized"

; Get current timestamp and username
FormatTime, currentTime,, yyyy-MM-dd HH:mm
currentUser := A_UserName

xlsxPath := A_ScriptDir . "\prompt_library.xlsx"

; Open Excel - attaches to the user's own open copy of prompt_library.xlsx
; if there is one, instead of spawning a second instance that would only
; be able to open the file read-only (see GetPromptWorkbook above).
xlHandle := ""
saveSucceeded := false
try {
    xlHandle := GetPromptWorkbook(xlsxPath)
    XL := xlHandle["XL"]
    wb := xlHandle["wb"]
    sheet := wb.Sheets(1)

    if (EditorMode = "new") {
        ; Find next empty row
        lastRow := sheet.Cells(sheet.Rows.Count, 2).End(-4162).Row
        newRow := lastRow + 1

        ; Write new prompt (format ID cell as text to preserve leading zeros)
        sheet.Cells(newRow, 1).NumberFormat := "@"
        ; Store Version, Created, CreatedBy, Modified, ModifiedBy (cols E-I) as text so
        ; Excel does not turn "1.0" into 1 or the timestamps into date serials
        sheet.Range(sheet.Cells(newRow, 5), sheet.Cells(newRow, 9)).NumberFormat := "@"
        sheet.Cells(newRow, 1).Value := newId
        sheet.Cells(newRow, 2).Value := newTitle
        sheet.Cells(newRow, 3).Value := newContent
        sheet.Cells(newRow, 4).Value := newCategory
        sheet.Cells(newRow, 5).Value := "1.0"           ; Version
        sheet.Cells(newRow, 6).Value := currentTime     ; Created
        sheet.Cells(newRow, 7).Value := currentUser     ; CreatedBy
        sheet.Cells(newRow, 8).Value := ""              ; Modified (empty for new)
        sheet.Cells(newRow, 9).Value := ""              ; ModifiedBy (empty for new)
        sheet.Cells(newRow, 10).Value := newStatus      ; Status
        sheet.Cells(newRow, 11).Value := newTags        ; Tags
        sheet.Cells(newRow, 12).Value := newDescription ; Description

        ; If the library is formatted as an Excel Table, extend it to include the new row
        try {
            if (sheet.ListObjects.Count > 0) {
                lo := sheet.ListObjects(1)
                lo.Resize(sheet.Range(sheet.Cells(1, 1), sheet.Cells(newRow, lo.Range.Columns.Count)))
            }
        }
    } else {
        ; Find the row to update by matching ID
        prompt := Prompts[EditingPromptIndex]
        oldId := prompt["id"]
        foundRow := 0

        lastRow := sheet.Cells(sheet.Rows.Count, 2).End(-4162).Row
        Loop, % (lastRow - 1) {
            row := A_Index + 1
            cellId := sheet.Cells(row, 1).Value
            if IsObject(cellId)
                cellId := cellId.Value
            cellId := cellId ""
            if (cellId = oldId) {
                foundRow := row
                break
            }
        }

        if (foundRow > 0) {
            ; Calculate new version (increment minor version)
            oldVersion := prompt["version"]
            if (oldVersion = "" or oldVersion = "0")
                oldVersion := "1.0"
            ; Parse version and increment minor
            versionParts := StrSplit(oldVersion, ".")
            majorVer := versionParts[1]
            minorVer := versionParts[2]
            if minorVer is not integer
                minorVer := 0
            minorVer += 1
            newVersion := majorVer . "." . minorVer

            ; Format ID cell as text to preserve leading zeros
            sheet.Cells(foundRow, 1).NumberFormat := "@"
            sheet.Range(sheet.Cells(foundRow, 5), sheet.Cells(foundRow, 9)).NumberFormat := "@"
            sheet.Cells(foundRow, 1).Value := newId
            sheet.Cells(foundRow, 2).Value := newTitle
            sheet.Cells(foundRow, 3).Value := newContent
            sheet.Cells(foundRow, 4).Value := newCategory
            sheet.Cells(foundRow, 5).Value := newVersion     ; Version (incremented)
            ; Preserve Created/CreatedBy (columns 6-7, don't overwrite)
            sheet.Cells(foundRow, 8).Value := currentTime    ; Modified
            sheet.Cells(foundRow, 9).Value := currentUser    ; ModifiedBy
            sheet.Cells(foundRow, 10).Value := newStatus     ; Status
            sheet.Cells(foundRow, 11).Value := newTags       ; Tags
            sheet.Cells(foundRow, 12).Value := newDescription ; Description
        } else {
            MsgBox, 16, Error, Could not find prompt row in Excel.
            ReleasePromptWorkbook(xlHandle)
            return
        }
    }

    ; Save and close (leaves the workbook open if it's the user's own window)
    wb.Save()
    ReleasePromptWorkbook(xlHandle)
    saveSucceeded := true

} catch e {
    errMsg := e.Message
    if InStr(errMsg, "read-only") or InStr(errMsg, "0x800A03EC") {
        MsgBox, 16, Error, prompt_library.xlsx couldn't be saved because it's open read-only somewhere else (for example, on a network share or by another user).`n`nClose it there, then try saving this prompt again.
    } else {
        MsgBox, 16, Error, % "Failed to save prompt:`n" . errMsg
    }
    ReleasePromptWorkbook(xlHandle)
}
xlHandle := ""
wb := ""
sheet := ""
XL := ""

; Only refresh if save succeeded
if (saveSucceeded) {
    ; Force reload of prompts
    CachedFileTime := ""
    PromptsLoaded := false

    ; Close editor
    Gui, 2:Destroy

    ; Reload prompts and refresh GUI
    Gosub, LoadPromptsIfNeeded
    Gosub, RefreshPromptList

    ToolTip, Prompt saved!
    SetTimer, RemoveTooltip, -1500
}
return

; Delete prompt from Excel
DeletePrompt:
MsgBox, 36, Confirm Delete, Are you sure you want to delete this prompt?`n`nThis cannot be undone.
IfMsgBox, No
    return

xlsxPath := A_ScriptDir . "\prompt_library.xlsx"
prompt := Prompts[EditingPromptIndex]
oldId := prompt["id"]

xlHandle := ""
try {
    xlHandle := GetPromptWorkbook(xlsxPath)
    XL := xlHandle["XL"]
    wb := xlHandle["wb"]
    sheet := wb.Sheets(1)

    ; Find the row to delete
    lastRow := sheet.Cells(sheet.Rows.Count, 2).End(-4162).Row
    foundRow := 0
    Loop, % (lastRow - 1) {
        row := A_Index + 1
        cellId := sheet.Cells(row, 1).Value
        if IsObject(cellId)
            cellId := cellId.Value
        cellId := cellId ""
        if (cellId = oldId) {
            foundRow := row
            break
        }
    }

    if (foundRow > 0) {
        ; Delete the entire row
        sheet.Rows(foundRow).Delete()

        wb.Save()
        ReleasePromptWorkbook(xlHandle)

        ; Force reload
        CachedFileTime := ""
        PromptsLoaded := false

        Gui, 2:Destroy

        Gosub, LoadPromptsIfNeeded
        Gosub, RefreshPromptList

        ToolTip, Prompt deleted!
        SetTimer, RemoveTooltip, -1500
    } else {
        MsgBox, 16, Error, Could not find prompt row in Excel.
        ReleasePromptWorkbook(xlHandle)
    }
} catch e {
    errMsg := e.Message
    if InStr(errMsg, "read-only") or InStr(errMsg, "0x800A03EC") {
        MsgBox, 16, Error, prompt_library.xlsx couldn't be saved because it's open read-only somewhere else (for example, on a network share or by another user).`n`nClose it there, then try deleting this prompt again.
    } else {
        MsgBox, 16, Error, % "Failed to delete prompt:`n" . errMsg
    }
    ReleasePromptWorkbook(xlHandle)
}
xlHandle := ""
wb := ""
sheet := ""
XL := ""
return

; Refresh the prompt list in main GUI after changes
RefreshPromptList:
; Rebuild category list string
CategoryListString := ""
Loop % Categories.Length() {
    if (A_Index > 1)
        CategoryListString .= "|"
    CategoryListString .= Categories[A_Index]
}
GuiControl, 1:, CategoryFilter, |%CategoryListString%
GuiControl, 1:Choose, CategoryFilter, 1

; Reset filters (category + tag search) to their defaults and rebuild list
; (Archived hidden by default here too, same rule as ApplyFilters' "All")
CurrentFilter := "All"
GuiControl, 1:, TagSearchBox,
Gosub, ApplyFilters
return

; ==================== SETTINGS DIALOG ====================

; Show settings dialog
ShowSettings:
NewHotkey := CurrentHotkey
NewHotkeyDisplay := HotkeyDisplay
IsRecording := false

Gui, Settings:Destroy
Gui, Settings:Font, s9, Segoe UI
Gui, Settings:Add, Text, x10 y10 w360 h20, Configure Hotkey for AHK Prompt Picker
Gui, Settings:Font, s8, Segoe UI
Gui, Settings:Add, Text, x10 y30 w360 h16 cGray, Created by Wesley Erickson
Gui, Settings:Font, s9, Segoe UI
Gui, Settings:Add, Text, x10 y54 w100 h20, Current Hotkey:
Gui, Settings:Add, Text, x110 y54 w180 h20 cGreen vCurrentHotkeyLabel, %HotkeyDisplay%

Gui, Settings:Add, GroupBox, x10 y84 w360 h100, Change Hotkey
Gui, Settings:Add, Text, x20 y109 w340 h20, Click the button, then press your key combo:
Gui, Settings:Add, Button, x20 y134 w120 h28 gRecordHotkey vRecordBtn, Record Hotkey
Gui, Settings:Add, Text, x150 y139 w130 h20 vNewHotkeyLabel, %HotkeyDisplay%

Gui, Settings:Add, Text, x10 y190 w360 h30 cGray, Tip: Use Ctrl, Alt, Shift, Win + a letter or F-key`nDefault: Ctrl+Alt+P

Gui, Settings:Add, GroupBox, x10 y226 w360 h104, Options
u_startupChecked := FileExist(StartupLink) ? 1 : 0
Gui, Settings:Add, CheckBox, x20 y248 w340 h20 vSettingsStartup Checked%u_startupChecked%, Start AHK Prompt Picker when I sign in to Windows
Gui, Settings:Add, Text, x20 y282 w235 h20, Estimated minutes saved per prompt use:
Gui, Settings:Add, Edit, x260 y279 w55 h22 Number vSettingsMinutes
Gui, Settings:Add, UpDown, Range1-240, %MinutesPerUse%
Gui, Settings:Font, s8, Segoe UI
Gui, Settings:Add, Text, x20 y306 w340 h16 cGray, Used for the time-saved estimate in Usage Stats.
Gui, Settings:Font, s9, Segoe UI

Gui, Settings:Add, Button, x10 y344 w60 h30 gSaveSettingsAndClose Default, Save
Gui, Settings:Add, Button, x75 y344 w60 h30 gCancelSettings, Cancel
Gui, Settings:Add, Button, x140 y344 w90 h30 gResetToDefault, Reset Default
Gui, Settings:Add, Button, x235 y344 w55 h30 gSettingsHelp, Help
Gui, Settings:Add, Button, x295 y344 w70 h30 gSettingsPrivacy, Privacy

Gui, Settings:Show, w380 h390, Settings
return

; Help button clicked from within Settings - opens the User Guide without closing Settings
SettingsHelp:
Gosub, ShowUserGuide
return

; Privacy button clicked from within Settings - opens the Privacy window without closing Settings
SettingsPrivacy:
Gosub, ShowPrivacy
return

; Record hotkey button clicked
RecordHotkey:
if (IsRecording)
    return

IsRecording := true
GuiControl, Settings:, RecordBtn, Press keys...
GuiControl, Settings:, NewHotkeyLabel, Waiting...

; Set a timer to capture the key press
SetTimer, CaptureHotkey, 50
return

; Capture the hotkey combination
CaptureHotkey:
r := CaptureComboPress(capHk, capDisp)
if (r = 0)
    return  ; nothing pressed yet, keep polling
SetTimer, CaptureHotkey, Off
IsRecording := false
GuiControl, Settings:, RecordBtn, Record Hotkey
if (r = 1) {
    NewHotkey := capHk
    NewHotkeyDisplay := capDisp
    GuiControl, Settings:, NewHotkeyLabel, %capDisp%
} else if (r = 2) {
    GuiControl, Settings:, NewHotkeyLabel, Need modifier!
} else {
    GuiControl, Settings:, NewHotkeyLabel, Cancelled
}
return

; Polls once for a hotkey combination. Shared by the Settings and first-run recorders.
; Returns 0 = nothing pressed yet, 1 = captured (hk and disp are set),
;         2 = main key pressed with no modifier, 3 = cancelled with Esc.
; Once a main key is seen, modifiers are collected while it is held down, so a small
; delay between pressing the modifier and the main key no longer causes "Need modifier!".
; Left, right, and middle mouse buttons are ignored so clicking the dialog cannot be captured.
CaptureComboPress(ByRef hk, ByRef disp) {
    mainKey := ""
    mainVk := ""
    Loop, 254 {
        if (A_Index = 1 or A_Index = 2 or A_Index = 4)
            continue
        vk := Format("vk{:02X}", A_Index)
        name := GetKeyName(vk)
        if (name = "" or name = "Control" or name = "Alt" or name = "Shift" or name = "LWin" or name = "RWin")
            continue
        if (name = "LControl" or name = "RControl" or name = "LAlt" or name = "RAlt" or name = "LShift" or name = "RShift")
            continue
        if GetKeyState(vk, "P") {
            mainKey := name
            mainVk := vk
            break
        }
    }
    if (mainKey = "")
        return 0

    if (mainKey = "Escape") {
        KeyWait, %mainVk%, T1
        return 3
    }

    ; Collect modifiers while the main key is held (2 second safety limit)
    hasCtrl := false
    hasAlt := false
    hasShift := false
    hasWin := false
    startTick := A_TickCount
    Loop {
        if GetKeyState("Ctrl", "P")
            hasCtrl := true
        if GetKeyState("Alt", "P")
            hasAlt := true
        if GetKeyState("Shift", "P")
            hasShift := true
        if (GetKeyState("LWin", "P") or GetKeyState("RWin", "P"))
            hasWin := true
        if (!GetKeyState(mainVk, "P") or (A_TickCount - startTick) > 2000)
            break
        Sleep, 10
    }

    mods := ""
    dmods := ""
    if (hasCtrl) {
        mods .= "^"
        dmods .= "Ctrl+"
    }
    if (hasAlt) {
        mods .= "!"
        dmods .= "Alt+"
    }
    if (hasShift) {
        mods .= "+"
        dmods .= "Shift+"
    }
    if (hasWin) {
        mods .= "#"
        dmods .= "Win+"
    }
    if (mods = "")
        return 2

    hk := mods . mainKey
    StringUpper, keyUpper, mainKey
    disp := dmods . keyUpper
    return 1
}

; Save settings and close dialog
SaveSettingsAndClose:
Gui, Settings:Submit, NoHide

; Apply options first: stats estimate and the Start with Windows shortcut
if SettingsMinutes is integer
{
    if (SettingsMinutes > 0)
        MinutesPerUse := SettingsMinutes
}
Gosub, SaveOptionSettings
u_hasLink := FileExist(StartupLink) ? true : false
if (SettingsStartup and !u_hasLink) {
    FileCreateShortcut, %A_ScriptFullPath%, %StartupLink%, %A_ScriptDir%, , AHK Prompt Picker
    if (ErrorLevel)
        MsgBox, 48, Start with Windows, The startup shortcut could not be created. Your computer's policy may block it.
} else if (!SettingsStartup and u_hasLink) {
    FileDelete, %StartupLink%
}

if (NewHotkey = "" or NewHotkey = CurrentHotkey) {
    if (NewHotkey = "") {
        MsgBox, 48, No Hotkey, Please record a hotkey first.
        return
    }
    ; No hotkey change, just close
    Gui, Settings:Destroy
    ToolTip, Settings saved
    SetTimer, RemoveTooltip, -1500
    return
}

; Try to register the new hotkey
; Disable old hotkey first (if one exists)
if (CurrentHotkey != "") {
    Hotkey, %CurrentHotkey%, Off, UseErrorLevel
    if (ErrorLevel) {
        MsgBox, 16, Error, Failed to disable current hotkey.
        return
    }
}

; Try to enable new hotkey
Hotkey, %NewHotkey%, TogglePromptPicker, On UseErrorLevel
if (ErrorLevel) {
    ; Failed - restore old hotkey if there was one
    if (CurrentHotkey != "")
        Hotkey, %CurrentHotkey%, On
    MsgBox, 16, Error, Failed to register hotkey: %NewHotkey%`n`nThe key combination may be invalid or already in use.`n`nPlease try a different combination.
    return
}

; Success - save settings
CurrentHotkey := NewHotkey
HotkeyDisplay := NewHotkeyDisplay
Gosub, SaveSettings

; Update tray tooltip
Menu, Tray, Tip, AHK Prompt Picker - %HotkeyDisplay%

Gui, Settings:Destroy
ToolTip, Hotkey changed to %HotkeyDisplay%!
SetTimer, RemoveTooltip, -2000
return

; Reset to default hotkey
ResetToDefault:
NewHotkey := "^!p"
NewHotkeyDisplay := "Ctrl+Alt+P"
GuiControl, Settings:, NewHotkeyLabel, %NewHotkeyDisplay%
return

; Cancel settings dialog
CancelSettings:
SettingsGuiClose:
SettingsGuiEscape:
SetTimer, CaptureHotkey, Off
IsRecording := false
Gui, Settings:Destroy
return

; ==================== END SETTINGS DIALOG ====================

; ==================== FIRST RUN SETUP ====================

; Show first-run setup dialog
ShowFirstRunSetup:
; Default to Ctrl+Alt+P - user can accept or record a custom one
NewHotkey := "^!p"
NewHotkeyDisplay := "Ctrl+Alt+P"
IsRecording := false

Gui, 6:Destroy
Gui, 6:Font, s11 Bold, Segoe UI
Gui, 6:Add, Text, x20 y12 w360 h25 +Center, Welcome to AHK Prompt Picker!
Gui, 6:Font, s9 Normal, Segoe UI
Gui, 6:Add, Text, x20 y38 w360 h16 +Center cGray, Created by Wesley Erickson
Gui, 6:Add, Text, x20 y60 w360 h34, Your hotkey is set to Ctrl+Alt+P by default. Keep it or record a different one, then we will walk you through how everything works.

Gui, 6:Add, GroupBox, x20 y95 w360 h100, Your Hotkey
Gui, 6:Add, Text, x35 y120 w330 h20, Or click to record a different key combination:
Gui, 6:Add, Button, x35 y145 w140 h28 gFirstRunRecord vFirstRunRecordBtn, Record Hotkey
Gui, 6:Add, Text, x185 y150 w180 h20 vFirstRunHotkeyLabel cGreen, Ctrl+Alt+P (Default)

Gui, 6:Add, Text, x20 y205 w360 h35 cGray, Tip: Use Ctrl, Alt, Shift, or Win + a letter or F-key.`nExample: Ctrl+Alt+Z, Ctrl+Shift+P, Win+F1

Gui, 6:Add, Button, x60 y250 w280 h35 gFirstRunSave Default, Finish Setup and Show Me Around

Gui, 6:Show, w400 h300, AHK Prompt Picker - First Time Setup
return

; Record hotkey in first-run dialog
FirstRunRecord:
if (IsRecording)
    return

IsRecording := true
GuiControl, 6:, FirstRunRecordBtn, Press keys...
GuiControl, 6:, FirstRunHotkeyLabel, Waiting...

SetTimer, CaptureFirstRunHotkey, 50
return

; Capture hotkey for first-run setup
CaptureFirstRunHotkey:
r := CaptureComboPress(capHk, capDisp)
if (r = 0)
    return  ; nothing pressed yet, keep polling
SetTimer, CaptureFirstRunHotkey, Off
IsRecording := false
GuiControl, 6:, FirstRunRecordBtn, Record Hotkey
if (r = 1) {
    NewHotkey := capHk
    NewHotkeyDisplay := capDisp
    GuiControl, 6:, FirstRunHotkeyLabel, %capDisp%
    ; Auto-advance after recording custom hotkey
    SetTimer, FirstRunAutoSave, -500
} else if (r = 2) {
    GuiControl, 6:, FirstRunHotkeyLabel, Need modifier!
} else {
    GuiControl, 6:, FirstRunHotkeyLabel, Cancelled
}
return

; Auto-save after recording (delayed to let user see their selection)
FirstRunAutoSave:
Gosub, FirstRunSave
return

; Save first-run hotkey and close
FirstRunSave:
if (NewHotkey = "") {
    MsgBox, 48, No Hotkey Set, Please record a hotkey before continuing.
    return
}

; Try to register the hotkey
Hotkey, %NewHotkey%, TogglePromptPicker, On UseErrorLevel
if (ErrorLevel) {
    MsgBox, 16, Error, Failed to register hotkey: %NewHotkeyDisplay%`n`nThe key combination may be invalid or already in use.`n`nPlease try a different combination.
    return
}

; Success - save settings
CurrentHotkey := NewHotkey
HotkeyDisplay := NewHotkeyDisplay
Gosub, SaveSettings

; Update tray tooltip
Menu, Tray, Tip, AHK Prompt Picker - %HotkeyDisplay%

Gui, 6:Destroy
FirstRun := false

; Launch the tailored onboarding guide. The picker is not opened here because
; the user may not have a cursor in the place they want to paste yet.
OnboardingMode := true
Gosub, ShowUserGuide
return

; Handle first-run dialog close
6GuiClose:
6GuiEscape:
SetTimer, CaptureFirstRunHotkey, Off
IsRecording := false
; If user closes without setting hotkey, exit the app
if (CurrentHotkey = "") {
    MsgBox, 48, Setup Required, You must set a hotkey to use AHK Prompt Picker.`n`nThe application will now exit.
    ExitApp
}
Gui, 6:Destroy
return

; ==================== END FIRST RUN SETUP ====================

; ==================== USER GUIDE DIALOG ====================

; Show user guide/help dialog. When OnboardingMode is true (first run) the
; guide opens with a personalized "get started" section and a different button.
ShowUserGuide:
Gui, 5:Destroy

; Size the window to the monitor work area so nothing is cut off; the text
; box scrolls when the content is taller than the window.
SysGet, guideWA, MonitorWorkArea
guideH := 600
guideAvail := guideWARight - guideWALeft
guideAvail := guideWABottom - guideWATop - 60
if (guideH > guideAvail)
    guideH := guideAvail
guideW := 470

if (OnboardingMode) {
    guideTitle := "Welcome to AHK Prompt Picker"
    guideSub := "Your setup is complete. Here is how to get started."
    guideBtn := "Got it, let's go"
    guideIntro =
(LTrim Join`r`n
YOU'RE ALL SET
- Your hotkey is %HotkeyDisplay%. It works in most apps.

TRY IT IN 3 STEPS
1. Click into any text field where you want a prompt (Copilot chat, Teams, Outlook, Word).
2. Press %HotkeyDisplay%. The picker opens on top of your window.
3. Select a prompt and press Enter. The prompt text is pasted at your cursor.

CAN'T FIND IT LATER?
- AHK Prompt Picker runs in the background. Look for its icon in the system tray near the clock (you may need to click the ^ arrow).
- Double-click the icon to open the picker, or right-click it for Settings and Help.
)
} else {
    guideTitle := "AHK Prompt Picker v2.1"
    guideSub := "Created by Wesley Erickson"
    guideBtn := "OK"
    guideIntro =
(LTrim Join`r`n
QUICK START
- Press %HotkeyDisplay% from most apps to open AHK Prompt Picker.
- Select a prompt and click Insert Prompt (or press Enter).
- The prompt text is pasted at your cursor position.
)
}

guideBody =
(LTrim Join`r`n
DYNAMIC VARIABLES
- Use {{VariableName}} in prompts for fill-in-the-blank fields.
- When inserting, you will be prompted to enter values.
- Example: {{MeetingName}}, {{Date}}, {{Attendees}}

FORMATTING TOOLBAR
- In the editor, use the toolbar buttons for markdown.
- Buttons: B (Bold), I (Italic), Code, Link, H1, H2, Bullets, Quotes

ORGANIZATION
- Filter by category using the dropdown.
- Type in the Search box to find prompts by title, tags, description, or content. Several words narrow the results.
- Star prompts as Favorites for quick access.
- Set Status to Archived to hide a prompt; use the Archived filter to bring it back.

FAVORITES, MOST USED, AND STATS
- Use the Filter by menu to show your Most used or Recently used prompts.
- Click Usage Stats to see how often you use each prompt and an estimate of time saved.
- Stats stay on your computer and record only prompt IDs and times, never prompt text.

TIPS
- Use Duplicate to create prompt variations.
- Hover over a prompt in the list to see its description.
- Add a Description to explain what a prompt is for.
- Change the hotkey any time in Settings (supports mouse buttons).
- In Settings, turn on the sign-in option to start automatically with Windows.

ADDING AND EDITING PROMPTS
- In the app: use New, Duplicate, or Edit below the prompt list.
- In Excel: click Open in Excel and type a new row directly below the last prompt. No blank rows are needed.
- Required: a unique ID, a Title, and the prompt Content.
- Blank Version, Category, and Status default to 1.0, Uncategorized, and Active.
- Your prompts live in prompt_library.xlsx, in the same folder as the script.
- Save the file. The picker loads your changes the next time it opens.

IF SOMETHING LOOKS OFF
- The first open can take a few seconds while your prompts load. You will see a "Loading your prompts" message, and the picker opens by itself.
- If the hotkey does nothing, look for the tray icon. If it is missing, run prompt-picker-p.ahk again.
- If another app already uses your hotkey, pick a different one in Settings.
- Click Privacy (or use the tray menu) to see what the app stores and how it handles your clipboard.
- If Open in Excel says the file is locked, close the other Excel window or user that has it open.
)

; The closing blank lines of a continuation section are dropped, so add the
; blank row between the intro and the body explicitly
guideText := guideIntro . "`r`n`r`n" . guideBody

Gui, 5:+Resize +MinSize380x300
Gui, 5:Font, s11 Bold, Segoe UI
Gui, 5:Add, Text, x15 y10 w440 h24 +Center vGuideTitleTxt, %guideTitle%
Gui, 5:Font, s9 Normal, Segoe UI
Gui, 5:Add, Text, x15 y36 w440 h18 +Center cGray vGuideSubTxt, %guideSub%
guideEditH := guideH - 58 - 56
g_rich := DllCall("LoadLibrary", "Str", "Msftedit.dll", "Ptr")
if (g_rich) {
    ; Rich text box so section headings can be bold
    Gui, 5:Add, Custom, ClassRichEdit50W x15 y58 w440 h%guideEditH% vGuideEdit hwndhGuideRE +0x200000 +0x4 +0x40 +E0x200 -Tabstop
    RichGuideSetup(hGuideRE, guideText)
} else {
    Gui, 5:Add, Edit, x15 y58 w440 h%guideEditH% vGuideEdit ReadOnly Multi +VScroll -Tabstop BackgroundWhite, %guideText%
}
guideBtnY := guideH - 44
Gui, 5:Add, Button, x170 y%guideBtnY% w130 h30 vGuideOkBtn gCloseUserGuide Default, %guideBtn%
Gui, 5:Show, w%guideW% h%guideH%, %guideTitle%
return

; Keep the guide layout correct when the window is resized
5GuiSize:
if (A_EventInfo = 1)
    return
guideNewEditW := A_GuiWidth - 30
guideNewEditH := A_GuiHeight - 58 - 56
guideNewBtnX := (A_GuiWidth - 130) // 2
guideNewBtnY := A_GuiHeight - 44
GuiControl, 5:Move, GuideTitleTxt, w%guideNewEditW%
GuiControl, 5:Move, GuideSubTxt, w%guideNewEditW%
GuiControl, 5:Move, GuideEdit, w%guideNewEditW% h%guideNewEditH%
GuiControl, 5:Move, GuideOkBtn, x%guideNewBtnX% y%guideNewBtnY%
return

; Close user guide dialog
CloseUserGuide:
5GuiClose:
5GuiEscape:
Gui, 5:Destroy
if (OnboardingMode) {
    OnboardingMode := false
    SetTimer, PrewarmPrompts, -500
    TrayTip, AHK Prompt Picker is ready, Click where you want to paste then press %HotkeyDisplay%, 8
}
return

; ==================== PRIVACY DIALOG ====================

; Show a plain-language privacy summary (scrollable, resizable, bold headings)
ShowPrivacy:
Gui, Privacy:Destroy
SysGet, pvWA, MonitorWorkArea
pvH := 560
pvAvail := pvWABottom - pvWATop - 60
if (pvH > pvAvail)
    pvH := pvAvail
pvW := 470

pvText =
(LTrim Join`r`n
PRIVACY SUMMARY
- AHK Prompt Picker runs entirely on your computer. It has no accounts, no analytics, and no tracking.
- It does not send your prompts, settings, or usage data anywhere. The app makes no network connections.

WHAT IT STORES
All files are kept in the same folder as the script.
- prompt_library.xlsx: your prompts, including titles, content, tags, and the Windows user name shown as author or editor.
- settings.ini: your hotkey and preferences.
- favorites.ini: the IDs of the prompts you starred.
- usage.ini: prompt IDs, use counts, and last-used times. It never records prompt text or the app you pasted into.
- A Startup shortcut, only if you turn on "Start when I sign in to Windows" in Settings.

CLIPBOARD AND WINDOWS
- To paste a prompt, the app briefly places it on the clipboard, pastes it, then restores what was there before.
- It remembers which window was active when you pressed the hotkey so it can paste back into it. This is kept in memory only.
- It does not read or record what you type or what is on your screen.

EXCEL
- The app reads and updates prompt_library.xlsx through Microsoft Excel on your computer. No cloud service is involved.

YOUR CONTROL
- You can view, edit, or delete any of these files at any time.
- Use Reset Stats in Usage Stats to clear your usage history.
- Deleting the app folder removes everything the app stores.

WORK AND ORGANIZATION DATA
- Prompts you paste go into the app you choose. Follow your organization's policies for sensitive or regulated data when you write prompts and use AI tools.
- The app reads prompt_library.xlsx from its own folder only. It does not connect to OneDrive, SharePoint, or any shared library.

CHANGES AND CONTACT
- This summary may change in future versions and is included with each release.
- Questions: pmpmaker.com
- This summary is provided for transparency and is not legal advice.
)

Gui, Privacy:+Resize +MinSize380x300
Gui, Privacy:Font, s11 Bold, Segoe UI
Gui, Privacy:Add, Text, x15 y10 w440 h24 +Center vPrivacyTitleTxt, Privacy
Gui, Privacy:Font, s9 Normal, Segoe UI
Gui, Privacy:Add, Text, x15 y36 w440 h18 +Center cGray vPrivacySubTxt, AHK Prompt Picker v2.1
pvEditH := pvH - 58 - 56
pvRich := DllCall("LoadLibrary", "Str", "Msftedit.dll", "Ptr")
if (pvRich) {
    Gui, Privacy:Add, Custom, ClassRichEdit50W x15 y58 w440 h%pvEditH% vPrivacyEdit hwndhPrivacyRE +0x200000 +0x4 +0x40 +E0x200 -Tabstop
    RichGuideSetup(hPrivacyRE, pvText)
} else {
    Gui, Privacy:Add, Edit, x15 y58 w440 h%pvEditH% vPrivacyEdit ReadOnly Multi +VScroll -Tabstop BackgroundWhite, %pvText%
}
pvBtnY := pvH - 44
Gui, Privacy:Add, Button, x170 y%pvBtnY% w130 h30 vPrivacyOkBtn gClosePrivacy Default, OK
Gui, Privacy:Show, w%pvW% h%pvH%, Privacy
return

PrivacyGuiSize:
if (A_EventInfo = 1)
    return
pvNewW := A_GuiWidth - 30
pvNewH := A_GuiHeight - 58 - 56
pvNewBtnX := (A_GuiWidth - 130) // 2
pvNewBtnY := A_GuiHeight - 44
GuiControl, Privacy:Move, PrivacyTitleTxt, w%pvNewW%
GuiControl, Privacy:Move, PrivacySubTxt, w%pvNewW%
GuiControl, Privacy:Move, PrivacyEdit, w%pvNewW% h%pvNewH%
GuiControl, Privacy:Move, PrivacyOkBtn, x%pvNewBtnX% y%pvNewBtnY%
return

ClosePrivacy:
PrivacyGuiClose:
PrivacyGuiEscape:
Gui, Privacy:Destroy
return

; ==================== END PRIVACY DIALOG ====================

; ==================== END USER GUIDE DIALOG ====================

; ==================== VARIABLE INPUT DIALOG ====================

; Show dialog for entering variable values
ShowVariableDialog:
Gui, 4:Destroy
Gui, 4:Font, s9, Segoe UI
Gui, 4:Add, Text, x10 y10 w350 h20, Fill in the values for this prompt:

; Calculate positions based on number of variables
yPos := 40
Loop % PromptVariables.Length() {
    varName := PromptVariables[A_Index]
    Gui, 4:Add, Text, x10 y%yPos% w100 h20, %varName%:
    editY := yPos - 3
    Gui, 4:Add, Edit, x120 y%editY% w240 h24 vVarInput%A_Index%
    yPos += 35
}

; Add buttons below the last variable
buttonY := yPos + 10
Gui, 4:Add, Button, x120 y%buttonY% w100 h30 gSubmitVariables Default, Insert
Gui, 4:Add, Button, x230 y%buttonY% w100 h30 gCancelVariables, Cancel

; Calculate dialog height
dialogHeight := buttonY + 50
Gui, 4:Show, w380 h%dialogHeight%, Enter Variable Values
return

; Submit variables and insert the prompt
SubmitVariables:
Gui, 4:Submit, NoHide

; Get values and substitute into content
substitutedContent := PendingPromptContent
Loop % PromptVariables.Length() {
    varName := PromptVariables[A_Index]
    GuiControlGet, varValue, 4:, VarInput%A_Index%
    ; Replace all occurrences of {{VarName}} with the value
    substitutedContent := StrReplace(substitutedContent, "{{" . varName . "}}", varValue)
}

; Close variable dialog
Gui, 4:Destroy

; Store substituted content and insert
PendingPromptContent := substitutedContent
Gosub, DoInsertContent
return

; Cancel variable dialog
CancelVariables:
4GuiClose:
4GuiEscape:
Gui, 4:Destroy
PendingPromptContent := ""
PromptVariables := []
return

; ==================== END VARIABLE INPUT DIALOG ====================

; Insert content directly where cursor was
InsertContent:
Loop % Prompts.Length() {
    if (Prompts[A_Index]["_display"] = CurrentPrompt) {
        prompt := Prompts[A_Index]
        content := prompt["content"]
        PendingPromptKey := UsageKey(prompt)

        ; Check for variables in the prompt
        PromptVariables := ExtractVariables(content)

        if (PromptVariables.Length() > 0) {
            ; Variables found - show dialog to fill them in
            PendingPromptContent := content
            Gosub, ShowVariableDialog
            return
        }

        ; No variables - insert directly
        PendingPromptContent := content
        Gosub, DoInsertContent
        break
    }
}
return

; Actual insert logic (called after variable substitution or directly)
DoInsertContent:
; Close the main GUI first (explicitly GUI 1)
Gui, 1:Destroy
GuiOpen := false

; Switch back to the original window
WinActivate, ahk_id %LastActiveWindow%
Sleep, 150  ; Give window more time to activate

; Use PendingPromptContent if set (from variable substitution), otherwise use prompt content
contentToInsert := (PendingPromptContent != "") ? PendingPromptContent : prompt["content"]

; Robust clipboard paste to emulate manual copy/paste exactly
ClipBackup := ClipboardAll        ; save full clipboard (binary-safe)
Clipboard := ""                ; clear clipboard
Sleep, 30
Clipboard := contentToInsert
; Wait for clipboard to contain the data (1 second timeout)
ClipWait, 1
if (ErrorLevel) {
    ; Clipboard not set in time - fallback to typing the content
    SendInput, %contentToInsert%
} else {
    Send, ^v
}
Sleep, 100
; Restore previous clipboard
if (ClipBackup != "")
    Clipboard := ClipBackup
else
    Clipboard := ""
ClipBackup := ""

; Clear pending content
PendingPromptContent := ""

; Record the use for the stats window (prompt ID and time only, never content)
RecordUsage(PendingPromptKey)
PendingPromptKey := ""

; Show brief confirmation
ToolTip, Content inserted!
SetTimer, RemoveTooltip, -1000
return

; Copy just the content to clipboard
CopyContent:
Loop % Prompts.Length() {
    if (Prompts[A_Index]["_display"] = CurrentPrompt) {
        prompt := Prompts[A_Index]
        ; Set clipboard to the raw prompt text from column C
        Clipboard := prompt["content"]
        RecordUsage(UsageKey(prompt))
        ToolTip, Prompt copied to clipboard!
        SetTimer, RemoveTooltip, -1500
        break
    }
}
return

; Helper function to escape JSON strings
EscapeJsonString(str) {
    str := StrReplace(str, "\", "\\")
    str := StrReplace(str, """", "\""")
    str := StrReplace(str, "`n", "\n")
    str := StrReplace(str, "`r", "\r")
    str := StrReplace(str, "`t", "\t")
    return str
}

; Extract unique {{VariableName}} placeholders from content
ExtractVariables(content) {
    variables := []
    seen := {}
    pos := 1
    while (pos := RegExMatch(content, "\{\{(\w+)\}\}", match, pos)) {
        varName := match1
        ; Only add if not already seen (handles duplicates)
        if (!seen[varName]) {
            variables.Push(varName)
            seen[varName] := true
        }
        pos += StrLen(match)
    }
    return variables
}

; Key used to track a prompt in usage.ini: its ID, or its title when the ID is
; blank. "=" is replaced because it is the ini key/value separator.
UsageKey(pp) {
    k := (pp["id"] != "") ? pp["id"] : pp["task"]
    return Trim(StrReplace(k, "=", "-"))
}

; Record one use of a prompt in usage.ini (prompt key, count, and time only)
RecordUsage(k) {
    global UsageCount, UsageLast, UsageFile
    if (k = "")
        return
    cnt := UsageCount.HasKey(k) ? UsageCount[k] : 0
    cnt += 1
    ts := A_Now
    UsageCount[k] := cnt
    UsageLast[k] := ts
    IniWrite, %cnt%, %UsageFile%, Uses, %k%
    IniWrite, %ts%, %UsageFile%, LastUsed, %k%
    IniRead, sn, %UsageFile%, Meta, Since, NONE
    if (sn = "NONE")
        IniWrite, %ts%, %UsageFile%, Meta, Since
}

; Fill a rich text box with the guide text and bold every ALL-CAPS heading line
RichGuideSetup(h, text) {
    ; Rich edit uses a single CR per paragraph, so offsets are line length + 1
    richText := StrReplace(text, "`r`n", "`r")
    DllCall("SendMessage", "Ptr", h, "UInt", 0x443, "Ptr", 0, "Ptr", 0xFFFFFF)       ; EM_SETBKGNDCOLOR white
    DllCall("SendMessage", "Ptr", h, "UInt", 0xC, "Ptr", 0, "Str", richText)         ; WM_SETTEXT

    ; Base font for all text: Segoe UI 9 pt
    VarSetCapacity(cf, 92, 0)
    NumPut(92, cf, 0, "UInt")
    NumPut(0xA0000000, cf, 4, "UInt")                                                ; CFM_FACE | CFM_SIZE
    NumPut(180, cf, 12, "Int")                                                       ; 9 pt in twips
    StrPut("Segoe UI", &cf + 26, 32, "UTF-16")
    DllCall("SendMessage", "Ptr", h, "UInt", 0x444, "Ptr", 4, "Ptr", &cf)            ; EM_SETCHARFORMAT, SCF_ALL

    ; Bold headings
    VarSetCapacity(bf, 92, 0)
    NumPut(92, bf, 0, "UInt")
    NumPut(1, bf, 4, "UInt")                                                         ; CFM_BOLD
    NumPut(1, bf, 8, "UInt")                                                         ; CFE_BOLD
    pos := 0
    Loop, Parse, richText, `r
    {
        line := A_LoopField
        len := StrLen(line)
        up := line
        StringUpper, up, up
        if (len > 3 and line == up and RegExMatch(line, "[A-Z]")) {
            DllCall("SendMessage", "Ptr", h, "UInt", 0xB1, "Ptr", pos, "Ptr", pos + len)   ; EM_SETSEL
            DllCall("SendMessage", "Ptr", h, "UInt", 0x444, "Ptr", 1, "Ptr", &bf)          ; SCF_SELECTION
        }
        pos += len + 1
    }
    DllCall("SendMessage", "Ptr", h, "UInt", 0xB1, "Ptr", 0, "Ptr", 0)               ; clear selection, caret at top
    DllCall("SendMessage", "Ptr", h, "UInt", 0xCF, "Ptr", 1, "Ptr", 0)               ; EM_SETREADONLY
}

; Word-wrap text for a tooltip (keeps long descriptions readable)
WrapTooltip(text, width := 60) {
    text := StrReplace(StrReplace(text, "`r`n", " "), "`n", " ")
    if (StrLen(text) > 500)
        text := SubStr(text, 1, 497) . "..."
    out := ""
    line := ""
    Loop, Parse, text, %A_Space%
    {
        word := A_LoopField
        if (word = "")
            continue
        if (line != "" and StrLen(line) + StrLen(word) + 1 > width) {
            out .= line . "`n"
            line := word
        } else {
            line := (line = "") ? word : line . " " . word
        }
    }
    return out . line
}

; Format a number of minutes as "X hr Y min"
FormatMinutes(m) {
    h := m // 60
    r := Mod(m, 60)
    if (h > 0)
        return h . " hr " . r . " min"
    return r . " min"
}

; Reorder a list of prompt indexes for the Most used / Recently used filters (stable, descending)
SortPromptIndexes(idxArr) {
    global Prompts, UsageCount, UsageLast, CurrentFilter
    if (CurrentFilter != "Most used" and CurrentFilter != "Recently used")
        return idxArr
    keys := []
    for n, pi in idxArr {
        pp := Prompts[pi]
        k := UsageKey(pp)
        if (CurrentFilter = "Most used")
            v := UsageCount.HasKey(k) ? UsageCount[k] : 0
        else
            v := UsageLast.HasKey(k) ? UsageLast[k] : 0
        keys.Push(v + 0)
    }
    Loop % idxArr.Length() - 1 {
        i := A_Index + 1
        curI := idxArr[i]
        curK := keys[i]
        j := i - 1
        while (j >= 1 and keys[j] < curK) {
            idxArr[j + 1] := idxArr[j]
            keys[j + 1] := keys[j]
            j -= 1
        }
        idxArr[j + 1] := curI
        keys[j + 1] := curK
    }
    return idxArr
}

; True when every word in the search text appears in the prompt's ID, title,
; category, tags, description, or content (case-insensitive)
PromptMatchesSearch(pp, text) {
    hay := pp["id"] . "`n" . pp["task"] . "`n" . pp["category"] . "`n" . pp["tags"] . "`n" . pp["description"] . "`n" . pp["content"]
    Loop, Parse, text, %A_Space%
    {
        if (A_LoopField = "")
            continue
        if !InStr(hay, A_LoopField)
            return false
    }
    return true
}

; Create a starter prompt_library.xlsx with a few example prompts (needs Excel)
CreateStarterLibrary(path) {
    global StarterLibError
    StarterLibError := ""
    ; Excel rejects doubled backslashes in SaveAs paths (keep a leading UNC prefix)
    path := SubStr(path, 1, 2) . StrReplace(SubStr(path, 3), "\\", "\")
    try {
        xl := ComObjCreate("Excel.Application")
        xl.Visible := false
        xl.DisplayAlerts := false
        wb := xl.Workbooks.Add()
        sheet := wb.Sheets(1)
        sheet.Name := "PromptLibrary"
        sheet.Columns(6).NumberFormat := "@"
        sheet.Columns(8).NumberFormat := "@"
        headers := ["ID", "Title", "Content", "Category", "Version", "Created", "CreatedBy", "Modified", "ModifiedBy", "Status", "Tags", "Description", "OutputFormat", "TargetApp"]
        for i, h in headers
            sheet.Cells(1, i).Value := h
        sheet.Rows(1).Font.Bold := true
        stamp := A_YYYY . "-" . A_MM . "-" . A_DD . " " . A_Hour . ":" . A_Min
        rows := []
        rows.Push(["GEN-001", "Summarize text", "Summarize the following text in 5 bullet points for a busy executive. Lead with the most important point.`n`nText: {{Text}}", "General", "1.0", "Summary, Writing", "Turns long text into five short bullets.", "Bullets", "CHAT"])
        rows.Push(["GEN-002", "Draft a professional email", "Draft a clear, professional email to {{Recipient}} about {{Topic}}. Keep it under 150 words and end with a specific next step.", "General", "1.0", "Email, Writing", "Quick first draft of a work email.", "Plain text", "OUTLOOK"])
        rows.Push(["GEN-003", "Explain in plain language", "Explain {{Topic}} in plain language for someone new to it. Use a short example and avoid unnecessary jargon.", "General", "1.0", "Learning, Explain", "Simple explanations of unfamiliar topics.", "Plain text", "CHAT"])
        ; Lighthearted examples: each one still does something useful
        rows.Push(["FUN-001", "Make no mistakes", "Complete the task below. Make no mistakes. I repeat: **no mistakes.**`n`nDouble-check your work because I have placed an unreasonable amount of confidence in this sentence.`n`nTask: {{Task}}", "Fun", "1.0", "Fun, Accuracy, Meme, AI", "The industry's most advanced quality-control framework.", "Plain text", "CHAT"])
        rows.Push(["FUN-002", "You are a 10x developer", "You are a legendary 10x developer who has seen things no ordinary engineer could comprehend. Solve the following problem elegantly, efficiently, and with suspiciously few bugs.`n`nProblem: {{Problem}}", "Fun", "1.0", "Fun, Coding, Meme, Vibe Coding", "Apparently multiplying the developer by ten improves the output.", "Plain text", "CHAT"])
        rows.Push(["FUN-003", "14 PhDs should be enough", "You are the world's foremost expert on {{Topic}}, with 40 years of experience, 14 PhDs, three Turing Awards, and somehow enough free time to answer my question.`n`nQuestion: {{Question}}", "Fun", "1.0", "Fun, Expert, Meme, AI", "Credentials generated faster than the answer.", "Plain text", "CHAT"])
        rows.Push(["FUN-004", "Don't hallucinate (please)", "Answer the following question accurately. Do not hallucinate. Please. Seriously. If you do not know something, say so instead of inventing a very convincing academic paper from 2017.`n`nQuestion: {{Question}}", "Fun", "1.0", "Fun, Accuracy, Meme, AI", "Asking politely has surely solved hallucinations by now.", "Plain text", "CHAT"])
        rows.Push(["FUN-005", "Fix it", "Something is broken. Fix it.`n`nProblem: {{Problem}}", "Fun", "1.0", "Fun, Coding, Debugging, Meme, Vibe Coding", "The complete bug report.", "Plain text", "CHAT"])
        rows.Push(["FUN-006", "Don't break anything", "Make the following change, but do not break anything else. This requirement is extremely important and therefore apparently sufficient testing.`n`nChange: {{Change}}", "Fun", "1.0", "Fun, Coding, Meme, Vibe Coding", "The traditional pre-refactor safety system.", "Plain text", "CHAT"])
        rows.Push(["FUN-007", "Make it production-ready", "Take this thing I generated five minutes ago and make it production-ready.`n`nThing: {{CodeOrIdea}}", "Fun", "1.0", "Fun, Coding, Production, Meme, Vibe Coding", "From prototype to Fortune 500 infrastructure in one prompt.", "Plain text", "CHAT"])
        rows.Push(["FUN-008", "Make it 10x better", "Make the following exactly 10x better. Not 9x. Not 11x. Use your judgment to determine the internationally recognized unit of ""better.""`n`nThing: {{Thing}}", "Fun", "1.0", "Fun, Improve, Meme, AI", "Quantitative improvement without quantitative requirements.", "Plain text", "CHAT"])
        rows.Push(["FUN-009", "I trust you. Figure it out.", "I am providing minimal context and maximum confidence.`n`nFigure out what needs to be done and do your best.`n`nTask: {{Task}}`n`nI trust you, which is probably irresponsible.", "Fun", "1.0", "Fun, Agent, Meme, Vibe Coding", "Requirements are temporary. Confidence is forever.", "Plain text", "CHAT"])
        rows.Push(["FUN-010", "Make the UI look modern", "Make this UI look modern.`n`nYou know the drill: tasteful spacing, cleaner typography, polished hierarchy, and enough rounded rectangles to confirm that it was designed after 2020.`n`nUI: {{UI}}", "Fun", "1.0", "Fun, UI, Design, Meme, Vibe Coding", "Add rounded corners until investors appear.", "Plain text", "CHAT"])
        for r, row in rows {
            n := r + 1
            sheet.Cells(n, 1).Value := row[1]
            sheet.Cells(n, 2).Value := row[2]
            sheet.Cells(n, 3).Value := row[3]
            sheet.Cells(n, 4).Value := row[4]
            sheet.Cells(n, 5).Value := row[5]
            sheet.Cells(n, 6).Value := stamp
            sheet.Cells(n, 7).Value := A_UserName
            sheet.Cells(n, 8).Value := stamp
            sheet.Cells(n, 9).Value := A_UserName
            sheet.Cells(n, 10).Value := "Active"
            sheet.Cells(n, 11).Value := row[6]
            sheet.Cells(n, 12).Value := row[7]
            sheet.Cells(n, 13).Value := row[8]
            sheet.Cells(n, 14).Value := row[9]
        }
        widths := [12, 40, 80, 25, 10, 18, 16, 18, 16, 10, 40, 60, 20, 20]
        for i, w in widths
            sheet.Columns(i).ColumnWidth := w
        wb.SaveAs(path, 51)
        wb.Close(false)
        xl.Quit()
        return true
    } catch e {
        StarterLibError := e.Message
        try xl.Quit()
        return false
    }
}

; GUI close handlers
ClosePromptPicker:
GuiClose:
GuiEscape:
Gui, 1:Destroy
GuiOpen := false
return

; Tooltip removal
RemoveTooltip:
ToolTip
return

; ==================== FORMATTING TOOLBAR ====================

; Clipboard-based wrap selection (more reliable in AHK)
; This mimics the user manually doing: Ctrl+C, modify, Ctrl+V
WrapSelectionClipboard(before, after, defaultText := "") {
    global

    ; Focus the editor first
    GuiControl, 2:Focus, EditorContent
    Sleep, 50

    ; Backup clipboard
    ClipBackup := ClipboardAll
    Clipboard := ""

    ; Try to copy current selection
    Send, ^c
    ClipWait, 0.3

    if (Clipboard != "") {
        ; Has selection - wrap it
        wrapped := before . Clipboard . after
        Clipboard := wrapped
        Send, ^v
    } else {
        ; No selection - insert default text
        Clipboard := before . defaultText . after
        Send, ^v
    }

    Sleep, 50

    ; Restore clipboard
    Clipboard := ClipBackup
    ClipBackup := ""
}

; Insert text at cursor position (for line prefixes)
InsertAtCursor(text) {
    global

    ; Focus the editor first
    GuiControl, 2:Focus, EditorContent
    Sleep, 50

    ; Backup clipboard
    ClipBackup := ClipboardAll
    Clipboard := text

    ; Paste the text
    Send, ^v
    Sleep, 50

    ; Restore clipboard
    Clipboard := ClipBackup
    ClipBackup := ""
}

; Add prefix at start of current line using Home key navigation
AddLinePrefixClipboard(prefix) {
    global

    ; Focus the editor
    GuiControl, 2:Focus, EditorContent
    Sleep, 50

    ; Save current position, go to line start, insert prefix, return
    Send, {Home}
    Sleep, 30

    ; Backup clipboard and insert prefix
    ClipBackup := ClipboardAll
    Clipboard := prefix
    Send, ^v
    Sleep, 30

    ; Move to end of line (where user probably wants to continue)
    Send, {End}

    ; Restore clipboard
    Clipboard := ClipBackup
    ClipBackup := ""
}

; ==================== FORMATTING BUTTON HANDLERS ====================

FmtBold:
WrapSelectionClipboard("**", "**", "bold text")
return

FmtItalic:
WrapSelectionClipboard("*", "*", "italic text")
return

FmtCode:
WrapSelectionClipboard("``", "``", "code")
return

FmtLink:
WrapSelectionClipboard("[", "](url)", "link text")
return

FmtH1:
AddLinePrefixClipboard("# ")
return

FmtH2:
AddLinePrefixClipboard("## ")
return

FmtBullet:
AddLinePrefixClipboard("- ")
return

FmtQuote:
AddLinePrefixClipboard("> ")
return

FmtCodeBlock:
WrapSelectionClipboard("```n", "`n```", "code")
return

FmtVar:
WrapSelectionClipboard("{{", "}}", "variable_name")
return

; ==================== END FORMATTING TOOLBAR ====================

; Disable double-click to insert (was causing issues)
; Optional: Double-click to insert (for power users)
; ~LButton::
; if (A_TimeSincePriorHotkey < 300 && A_PriorHotkey = "~LButton") {
;     Gosub, InsertContent
; }
; return