//go:build windows

package main

// A single-window, native Win32 front end with NO additional runtime/library
// dependencies. The path field + action button are centered vertically;
// the progress bar/status or a result label occupy the bottom row.

import (
	"context"
	"fmt"
	"os"
	"path/filepath"
	"runtime"
	"strings"
	"sync"
	"syscall"
	"unsafe"
)

const (
	csHRedraw        = 0x0002
	csVRedraw        = 0x0001
	wsVisible        = 0x10000000
	wsChild          = 0x40000000
	wsTabStop        = 0x00010000
	wsBorder         = 0x00800000
	wsCaption        = 0x00C00000
	wsSysmenu        = 0x00080000
	wsMinimizeBox    = 0x00020000
	wsClipChildren   = 0x02000000
	esAutoHscroll    = 0x0080
	wmCreate         = 0x0001
	wmDestroy        = 0x0002
	wmClose          = 0x0010
	wmSize           = 0x0005
	wmCommand        = 0x0111
	wmSetFont        = 0x0030
	wmGetText        = 0x000D
	emSetCueBanner   = 0x1501
	enChange         = 0x0300
	bnClicked        = 0x0000
	wmAppProgress    = 0x8001
	wmAppComplete    = 0x8002
	wmUser           = 0x0400
	pbmSetRange32    = wmUser + 6
	pbmSetPos        = wmUser + 2
	bsPushButton     = 0x0000
	swShow           = 5
	swHide           = 0
	mbIconError      = 0x00000010
	mbIconInfo       = 0x00000040
	iccProgressClass = 0x00000020
	ssRight          = 0x00000002
	ssCenter         = 0x00000001
	idEdit           = 101
	idButton         = 102
	idProgress       = 103
	idStatus         = 104
	idResult         = 105
)

type point struct{ X, Y int32 }
type rect struct{ Left, Top, Right, Bottom int32 }
type wndClassEx struct {
	Size       uint32
	Style      uint32
	WndProc    uintptr
	ClsExtra   int32
	WndExtra   int32
	Instance   uintptr
	Icon       uintptr
	Cursor     uintptr
	Background uintptr
	MenuName   *uint16
	ClassName  *uint16
	IconSmall  uintptr
}
type winMessage struct {
	Hwnd     uintptr
	Message  uint32
	WParam   uintptr
	LParam   uintptr
	Time     uint32
	Pt       point
	LPrivate uint32
}
type browseInfo struct {
	Owner       uintptr
	Root        uintptr
	DisplayName *uint16
	Title       *uint16
	Flags       uint32
	Callback    uintptr
	LParam      uintptr
	Image       int32
}
type commonControls struct {
	Size    uint32
	Classes uint32
}

var (
	dllUser              = syscall.NewLazyDLL("user32.dll")
	dllKernel            = syscall.NewLazyDLL("kernel32.dll")
	dllGdi               = syscall.NewLazyDLL("gdi32.dll")
	dllShell             = syscall.NewLazyDLL("shell32.dll")
	dllOle               = syscall.NewLazyDLL("ole32.dll")
	dllComCtl            = syscall.NewLazyDLL("comctl32.dll")
	procRegisterClass    = dllUser.NewProc("RegisterClassExW")
	procCreateWindow     = dllUser.NewProc("CreateWindowExW")
	procDefWindow        = dllUser.NewProc("DefWindowProcW")
	procGetMessage       = dllUser.NewProc("GetMessageW")
	procTranslateMessage = dllUser.NewProc("TranslateMessage")
	procDispatchMessage  = dllUser.NewProc("DispatchMessageW")
	procPostQuit         = dllUser.NewProc("PostQuitMessage")
	procPostMessage      = dllUser.NewProc("PostMessageW")
	procSendMessage      = dllUser.NewProc("SendMessageW")
	procSetText          = dllUser.NewProc("SetWindowTextW")
	procGetText          = dllUser.NewProc("GetWindowTextW")
	procEnableWindow     = dllUser.NewProc("EnableWindow")
	procMoveWindow       = dllUser.NewProc("MoveWindow")
	procShowWindow       = dllUser.NewProc("ShowWindow")
	procUpdateWindow     = dllUser.NewProc("UpdateWindow")
	procGetClientRect    = dllUser.NewProc("GetClientRect")
	procGetWindowRect    = dllUser.NewProc("GetWindowRect")
	procClientToScreen   = dllUser.NewProc("ClientToScreen")
	procGetSysMetrics    = dllUser.NewProc("GetSystemMetrics")
	procAdjustWindowRect = dllUser.NewProc("AdjustWindowRectEx")
	procLoadCursor       = dllUser.NewProc("LoadCursorW")
	procMessageBox       = dllUser.NewProc("MessageBoxW")
	procCreateFont       = dllGdi.NewProc("CreateFontW")
	procDeleteObject     = dllGdi.NewProc("DeleteObject")
	procGetModule        = dllKernel.NewProc("GetModuleHandleW")
	procInitCommon       = dllComCtl.NewProc("InitCommonControlsEx")
	procBrowseFolder     = dllShell.NewProc("SHBrowseForFolderW")
	procGetFolderPath    = dllShell.NewProc("SHGetPathFromIDListW")
	procCoTaskMemFree    = dllOle.NewProc("CoTaskMemFree")
	procCoInitialize     = dllOle.NewProc("CoInitializeEx")
	procCoUninitialize   = dllOle.NewProc("CoUninitialize")
)

type appUI struct {
	hwnd, edit, button, bar, status, resultLabel, font uintptr
	mu                                                 sync.Mutex
	updating                                           bool
	running                                            bool
	canceling                                          bool
	cancel                                             context.CancelFunc
	completed                                          int
	total                                              int
	result                                             batchResult
	finished                                           bool
	closeRequested                                     bool
}

var app = &appUI{}
var moduleInstance uintptr

func utfPtr(s string) *uint16 { return syscall.StringToUTF16Ptr(s) }
func send(hwnd uintptr, msg uint32, w, l uintptr) uintptr {
	r, _, _ := procSendMessage.Call(hwnd, uintptr(msg), w, l)
	return r
}
func setText(hwnd uintptr, s string) { procSetText.Call(hwnd, uintptr(unsafe.Pointer(utfPtr(s)))) }
func getText(hwnd uintptr) string {
	buf := make([]uint16, 32768)
	n, _, _ := procGetText.Call(hwnd, uintptr(unsafe.Pointer(&buf[0])), uintptr(len(buf)))
	return syscall.UTF16ToString(buf[:int(n)])
}
func create(class, name string, style uint32, x, y, width, height int32, parent uintptr, id uintptr) uintptr {
	hwnd, _, _ := procCreateWindow.Call(0, uintptr(unsafe.Pointer(utfPtr(class))), uintptr(unsafe.Pointer(utfPtr(name))), uintptr(style), uintptr(x), uintptr(y), uintptr(width), uintptr(height), parent, id, moduleInstance, 0)
	return hwnd
}
func msgBox(hwnd uintptr, text, caption string, flags uintptr) {
	procMessageBox.Call(hwnd, uintptr(unsafe.Pointer(utfPtr(text))), uintptr(unsafe.Pointer(utfPtr(caption))), flags)
}
func enable(hwnd uintptr, value bool) {
	n := uintptr(0)
	if value {
		n = 1
	}
	procEnableWindow.Call(hwnd, n)
}
func move(hwnd uintptr, x, y, w, h int32) {
	procMoveWindow.Call(hwnd, uintptr(x), uintptr(y), uintptr(w), uintptr(h), 1)
}
func show(hwnd uintptr, visible bool) {
	cmd := uintptr(swHide)
	if visible {
		cmd = swShow
	}
	procShowWindow.Call(hwnd, cmd)
}
func setFont(hwnd uintptr) { send(hwnd, wmSetFont, app.font, 1) }
func digits(n int) int {
	if n < 0 {
		n = -n
	}
	if n < 10 {
		return 1
	}
	d := 0
	for n > 0 {
		d++
		n /= 10
	}
	return d
}
func statusWidthForTotal(total int) int32 {
	// "Converting: X of Y" width depends mainly on the number of digits in Y.
	charCount := len("Converting: ") + digits(total) + len(" of ") + digits(total)
	return int32(24 + charCount*6)
}
func clearResult() {
	setText(app.resultLabel, "")
	show(app.resultLabel, false)
}
func setProgress(completed, total int) {
	if total <= 0 {
		total = 1
	}
	send(app.bar, pbmSetRange32, 0, uintptr(total))
	send(app.bar, pbmSetPos, uintptr(completed), 0)
}

func layout() {
	if app.edit == 0 {
		return
	}
	var r rect
	procGetClientRect.Call(app.hwnd, uintptr(unsafe.Pointer(&r)))
	w, h := r.Right-r.Left, r.Bottom-r.Top
	const margin int32 = 18
	const gap int32 = 12
	const rowH int32 = 28
	const buttonW int32 = 112
	const bottomGap int32 = 8
	const statusGap int32 = 10
	// Align the input/button row to the center of the ENTIRE window,
	// including its native title bar (which changes with DPI and theme).
	// Client coordinates alone would position it too low visually.
	nonClientOffset := int32(12) // fallback if native frame metrics fail
	var outer rect
	var topLeft point
	if ok, _, _ := procGetWindowRect.Call(app.hwnd, uintptr(unsafe.Pointer(&outer))); ok != 0 {
		if ok, _, _ := procClientToScreen.Call(app.hwnd, uintptr(unsafe.Pointer(&topLeft))); ok != 0 {
			ncTop := topLeft.Y - outer.Top
			ncBottom := outer.Bottom - (topLeft.Y + h)
			nonClientOffset = (ncTop - ncBottom) / 2
		}
	}
	y := (h-rowH)/2 - nonClientOffset
	if y < margin {
		y = margin
	}
	editW := w - 2*margin - buttonW - gap
	if editW < 120 {
		editW = 120
	}
	move(app.edit, margin, y, editW, rowH)
	move(app.button, margin+editW+gap, y, buttonW, rowH)

	statusW := statusWidthForTotal(app.total)
	statusMinW := int32(120)
	if statusW < statusMinW {
		statusW = statusMinW
	}
	maxStatusW := w - 2*margin - 120
	if statusW > maxStatusW {
		statusW = maxStatusW
	}
	statusX := w - margin - statusW
	barX := margin
	barY := h - 28 - bottomGap
	barW := statusX - statusGap - barX
	if barW < 120 {
		barW = 120
	}
	move(app.bar, barX, barY, barW, 18)
	move(app.status, statusX, barY-3, statusW, 24)
	// Centered result message, shown only outside conversion.
	move(app.resultLabel, margin, barY-3, w-2*margin, 24)
}

func folderPicker(owner uintptr) string {
	display := make([]uint16, 260)
	bi := browseInfo{Owner: owner, DisplayName: &display[0], Title: utfPtr("Select a folder containing DOCX files"), Flags: 0x0001 | 0x0040 | 0x0010}
	pidl, _, _ := procBrowseFolder.Call(uintptr(unsafe.Pointer(&bi)))
	if pidl == 0 {
		return ""
	}
	defer procCoTaskMemFree.Call(pidl)
	buf := make([]uint16, 32768)
	ok, _, _ := procGetFolderPath.Call(pidl, uintptr(unsafe.Pointer(&buf[0])))
	if ok == 0 {
		return ""
	}
	return syscall.UTF16ToString(buf)
}

func actionFromPath() {
	if app.updating || app.running {
		return
	}
	path := strings.TrimSpace(getText(app.edit))
	label := "Browse"
	result := "Select a directory"
	if path != "" && filepath.IsAbs(path) {
		if info, err := os.Stat(path); err == nil && info.IsDir() {
			label = "Convert"
			result = "Ready to convert"
		}
	}
	setText(app.button, label)
	setText(app.resultLabel, result)
	show(app.resultLabel, true)
}

func beginConversion() {
	path := strings.TrimSpace(getText(app.edit))
	docs, err := collectDocs(path)
	if err != nil {
		msgBox(app.hwnd, "Please choose an existing folder.", "DOCX to PDF", mbIconError)
		return
	}
	if len(docs) == 0 {
		msgBox(app.hwnd, "The selected folder contains no .docx files.", "DOCX to PDF", mbIconInfo)
		return
	}
	ctx, cancel := context.WithCancel(context.Background())
	app.mu.Lock()
	app.running = true
	app.canceling = false
	app.finished = false
	app.cancel = cancel
	app.completed = 0
	app.total = len(docs)
	app.mu.Unlock()
	layout()
	clearResult()
	show(app.bar, true)
	show(app.status, true)
	app.updating = true
	setText(app.edit, "")
	app.updating = false
	enable(app.edit, false)
	setText(app.button, "Cancel")
	setText(app.status, fmt.Sprintf("Converting: 0 of %d", len(docs)))
	setProgress(0, len(docs))
	go func() {
		res := convertBatch(ctx, path, func(done, total int) {
			app.mu.Lock()
			app.completed = done
			app.total = total
			app.mu.Unlock()
			procPostMessage.Call(app.hwnd, wmAppProgress, 0, 0)
		})
		app.mu.Lock()
		app.result = res
		app.finished = true
		app.mu.Unlock()
		procPostMessage.Call(app.hwnd, wmAppComplete, 0, 0)
	}()
}

func requestCancel() {
	app.mu.Lock()
	if !app.running || app.canceling {
		app.mu.Unlock()
		return
	}
	if app.finished {
		app.mu.Unlock()
		finishConversion()
		return
	}
	app.canceling = true
	cancel := app.cancel
	app.mu.Unlock()
	enable(app.button, false)
	setText(app.status, "Canceling...")
	cancel()
}

func finishConversion() {
	app.mu.Lock()
	res := app.result
	app.running = false
	app.canceling = false
	app.finished = false
	closeRequested := app.closeRequested
	app.closeRequested = false
	app.cancel = nil
	app.total = 0
	app.mu.Unlock()
	enable(app.edit, true)
	enable(app.button, true)
	setText(app.button, "Browse")
	show(app.bar, false)
	show(app.status, false)
	resultText := ""
	if res.canceled {
		setProgress(0, 1)
		resultText = "Canceled"
	} else if res.err != nil {
		resultText = "Failed"
		msgBox(app.hwnd, res.err.Error(), "DOCX to PDF", mbIconError)
	} else {
		setProgress(res.converted, res.total)
		if res.failed > 0 {
			resultText = fmt.Sprintf("Done: %d converted, %d failed", res.converted, res.failed)
			msgBox(app.hwnd, fmt.Sprintf("%d documents could not be converted.", res.failed), "DOCX to PDF", mbIconInfo)
		} else {
			resultText = fmt.Sprintf("Done: %d of %d", res.converted, res.total)
		}
	}
	setText(app.resultLabel, resultText)
	show(app.resultLabel, resultText != "")
	layout()
	if closeRequested {
		dllUser.NewProc("DestroyWindow").Call(app.hwnd)
	}
}

func windowProc(hwnd uintptr, message, wParam, lParam uintptr) uintptr {
	switch uint32(message) {
	case wmCreate:
		app.hwnd = hwnd
		app.edit = create("EDIT", "", wsVisible|wsChild|wsTabStop|wsBorder|esAutoHscroll, 0, 0, 0, 0, hwnd, idEdit)
		app.button = create("BUTTON", "Browse", wsVisible|wsChild|wsTabStop|bsPushButton, 0, 0, 0, 0, hwnd, idButton)
		app.bar = create("msctls_progress32", "", wsVisible|wsChild, 0, 0, 0, 0, hwnd, idProgress)
		app.status = create("STATIC", "", wsVisible|wsChild|ssRight, 0, 0, 0, 0, hwnd, idStatus)
		app.resultLabel = create("STATIC", "", wsVisible|wsChild|ssCenter, 0, 0, 0, 0, hwnd, idResult)
		for _, h := range []uintptr{app.edit, app.button, app.status, app.resultLabel} {
			setFont(h)
		}
		send(app.edit, emSetCueBanner, 0, uintptr(unsafe.Pointer(utfPtr("Select a folder with .docx files"))))
		setProgress(0, 1)
		show(app.bar, false)
		show(app.status, false)
		setText(app.resultLabel, "Select a directory")
		show(app.resultLabel, true)
		layout()
		return 0
	case wmSize:
		layout()
		return 0
	case wmCommand:
		id := uint16(wParam & 0xffff)
		notify := uint16((wParam >> 16) & 0xffff)
		if id == idEdit && notify == enChange {
			actionFromPath()
			return 0
		}
		if id == idButton && notify == bnClicked {
			app.mu.Lock()
			running := app.running
			app.mu.Unlock()
			if running {
				requestCancel()
				return 0
			}
			if getText(app.button) == "Convert" {
				beginConversion()
			} else {
				dir := folderPicker(hwnd)
				if dir != "" {
					setText(app.edit, dir)
					actionFromPath()
				}
			}
			return 0
		}
	case wmAppProgress:
		app.mu.Lock()
		done, total, canceling := app.completed, app.total, app.canceling
		app.mu.Unlock()
		if !canceling {
			layout()
			setProgress(done, total)
			setText(app.status, fmt.Sprintf("Converting: %d of %d", done, total))
		}
		return 0
	case wmAppComplete:
		finishConversion()
		return 0
	case wmClose:
		app.mu.Lock()
		running := app.running
		app.mu.Unlock()
		if running {
			app.mu.Lock()
			app.closeRequested = true
			app.mu.Unlock()
			requestCancel()
			return 0
		}
		// DestroyWindow only after any running batch and printer restoration finish.
		dllUser.NewProc("DestroyWindow").Call(hwnd)
		return 0
	case wmDestroy:
		if app.font != 0 {
			procDeleteObject.Call(app.font)
		}
		procPostQuit.Call(0)
		return 0
	}
	r, _, _ := procDefWindow.Call(hwnd, uintptr(message), wParam, lParam)
	return r
}

func main() {
	runtime.LockOSThread()
	defer runtime.UnlockOSThread()
	// Native folder picker requires an initialized STA on the UI thread.
	if hr, _, _ := procCoInitialize.Call(0, 2); int32(hr) >= 0 {
		defer procCoUninitialize.Call()
	}
	// The default manifest does not include DPI awareness; use system DPI.
	if p := dllUser.NewProc("SetProcessDPIAware"); p.Find() == nil {
		p.Call()
	}
	ctrl := commonControls{Size: uint32(unsafe.Sizeof(commonControls{})), Classes: iccProgressClass}
	procInitCommon.Call(uintptr(unsafe.Pointer(&ctrl)))
	inst, _, _ := procGetModule.Call(0)
	moduleInstance = inst
	className := utfPtr("DOCXtoPDFNativeGUI")
	callback := syscall.NewCallback(windowProc)
	cursor, _, _ := procLoadCursor.Call(0, 32512) // IDC_ARROW
	wc := wndClassEx{Size: uint32(unsafe.Sizeof(wndClassEx{})), Style: csHRedraw | csVRedraw, WndProc: callback, Instance: inst, Cursor: cursor, Background: 6, ClassName: className}
	if ret, _, _ := procRegisterClass.Call(uintptr(unsafe.Pointer(&wc))); ret == 0 {
		msgBox(0, "Unable to initialize window.", "DOCX to PDF", mbIconError)
		return
	}
	// Standard Windows chrome, minimize and close controls, not a custom frame.
	style := uint32(wsCaption | wsSysmenu | wsMinimizeBox | wsClipChildren)
	// Keep width; reduce client height from 180 to 136 (-24.4%).
	r := rect{Right: 480, Bottom: 136}
	procAdjustWindowRect.Call(uintptr(unsafe.Pointer(&r)), uintptr(style), 0)
	width, height := r.Right-r.Left, r.Bottom-r.Top
	screenW, _, _ := procGetSysMetrics.Call(0)
	screenH, _, _ := procGetSysMetrics.Call(1)
	x := (int32(screenW) - width) / 2
	y := (int32(screenH) - height) / 2
	// Smaller Segoe UI, close to compact Explorer sidebar text.
	app.font, _, _ = procCreateFont.Call(^uintptr(12), 0, 0, 0, 400, 0, 0, 0, 1, 0, 0, 0, 0, uintptr(unsafe.Pointer(utfPtr("Segoe UI"))))
	hwnd := create("DOCXtoPDFNativeGUI", "DOCX to PDF", style, x, y, width, height, 0, 0)
	if hwnd == 0 {
		msgBox(0, "Unable to create the application window.", "DOCX to PDF", mbIconError)
		return
	}
	procShowWindow.Call(hwnd, swShow)
	procUpdateWindow.Call(hwnd)
	var msg winMessage
	for {
		result, _, _ := procGetMessage.Call(uintptr(unsafe.Pointer(&msg)), 0, 0, 0)
		if result == 0 || result == ^uintptr(0) {
			break
		}
		procTranslateMessage.Call(uintptr(unsafe.Pointer(&msg)))
		procDispatchMessage.Call(uintptr(unsafe.Pointer(&msg)))
	}
}

// Word.Application.Hwnd identifies the isolated COM instance, so Cancel can
// terminate the correct Word process without killing Word windows opened by the user.
func wordPID(hwnd uintptr) int {
	if hwnd == 0 {
		return 0
	}
	var pid uint32
	dllUser.NewProc("GetWindowThreadProcessId").Call(hwnd, uintptr(unsafe.Pointer(&pid)))
	return int(pid)
}
