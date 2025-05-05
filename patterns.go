package uiautomation

import (
	"syscall"
	"unsafe"

	"github.com/go-ole/go-ole"
)

type PatternId uintptr

const (
	AnnotationPatternId        PatternId = 10023
	DockPatternId              PatternId = 10011
	DragPatternId              PatternId = 10030
	DropTargetPatternId        PatternId = 10031
	ExpandCollapsePatternId    PatternId = 10005
	GridItemPatternId          PatternId = 10007
	GridPatternId              PatternId = 10006
	InvokePatternId            PatternId = 10000
	ItemContainerPatternId     PatternId = 10019
	LegacyIAccessiblePatternId PatternId = 10018
	MultipleViewPatternId      PatternId = 10008
	ObjectModelPatternId       PatternId = 10022
	RangeValuePatternId        PatternId = 10003
	ScrollItemPatternId        PatternId = 10017
	ScrollPatternId            PatternId = 10004
	SelectionItemPatternId     PatternId = 10010
	SelectionPatternId         PatternId = 10001
	SpreadsheetPatternId       PatternId = 10026
	SpreadsheetItemPatternId   PatternId = 10027
	StylesPatternId            PatternId = 10025
	SynchronizedInputPatternId PatternId = 10021
	TableItemPatternId         PatternId = 10013
	TablePatternId             PatternId = 10012
	TextChildPatternId         PatternId = 10029
	TextEditPatternId          PatternId = 10032
	TextPatternId              PatternId = 10014
	TextPattern2Id             PatternId = 10024
	TogglePatternId            PatternId = 10015
	TransformPatternId         PatternId = 10016
	TransformPattern2Id        PatternId = 10028
	ValuePatternId             PatternId = 10002
	VirtualizedItemPatternId   PatternId = 10020
	WindowPatternId            PatternId = 10009
)

type UIAutomationTextPattern struct {
	ole.IUnknown
}

type UIAutomationTextPatternVtbl struct {
	ole.IUnknownVtbl
	RangeFromPoint             uintptr
	RangeFromChild             uintptr
	GetSelection               uintptr
	GetVisibleRanges           uintptr
	Get_DocumentRange          uintptr
	Get_SupportedTextSelection uintptr
}

func (pat *UIAutomationTextPattern) VTable() *UIAutomationTextPatternVtbl {
	return (*UIAutomationTextPatternVtbl)(unsafe.Pointer(pat.RawVTable))
}

func (pat *UIAutomationTextPattern) GetSelection() (*UIAutomationTextRangeArray, error) {
	var retVal *UIAutomationTextRangeArray

	hr, _, _ := syscall.SyscallN(
		pat.VTable().GetSelection,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&retVal)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return retVal, nil
}

type UIAutomationValuePattern struct {
	ole.IUnknown
}

type UIAutomationValuePatternVtbl struct {
	ole.IUnknownVtbl
	SetValue                   uintptr
	Get_CurrentValue           uintptr
	Get_CurrentIsReadonly      uintptr
	Get_CachedValue            uintptr
	Get_CachedCachedIsReadOnly uintptr
}

func (pat *UIAutomationValuePattern) VTable() *UIAutomationValuePatternVtbl {
	return (*UIAutomationValuePatternVtbl)(unsafe.Pointer(pat.RawVTable))
}

func (pat *UIAutomationValuePattern) SetValue(value string) error {
	valuePtr, err := syscall.UTF16PtrFromString(value)

	if err != nil {
		return err
	}

	hr, _, _ := syscall.SyscallN(
		pat.VTable().SetValue,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(valuePtr)),
	)

	if hr != 0 {
		return ole.NewError(hr)
	}

	return nil
}

func (pat *UIAutomationValuePattern) CurrentValue() (string, error) {
	var value *uint16

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CurrentValue,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&value)),
	)

	if hr != 0 {
		return "", ole.NewError(hr)
	}

	return ole.BstrToString(value), nil
}

func (pat *UIAutomationValuePattern) CurrentIsReadonly() (bool, error) {
	var readonly bool

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CurrentIsReadonly,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&readonly)),
	)

	if hr != 0 {
		return false, ole.NewError(hr)
	}

	return readonly, nil
}

type UIAutomationInvokePattern struct {
	ole.IUnknown
}

type UIAutomationInvokePatternVtbl struct {
	ole.IUnknownVtbl
	Invoke uintptr
}

func (pat *UIAutomationInvokePattern) VTable() *UIAutomationInvokePatternVtbl {
	return (*UIAutomationInvokePatternVtbl)(unsafe.Pointer(pat.RawVTable))
}

func (pat *UIAutomationInvokePattern) Invoke() error {
	hr, _, _ := syscall.SyscallN(
		pat.VTable().Invoke,
		uintptr(unsafe.Pointer(pat)),
	)

	if hr != 0 {
		return ole.NewError(hr)
	}

	return nil
}

type UIAutomationWindowPattern struct {
	ole.IUnknown
}

type UIAutomationWindowPatternVtbl struct {
	ole.IUnknownVtbl
	Close                             uintptr
	WaitForInputIdle                  uintptr
	SetWindowVisualState              uintptr
	Get_CurrentCanMaximize            uintptr
	Get_CurrentCanMinimize            uintptr
	Get_CurrentIsModal                uintptr
	Get_CurrentIsTopmost              uintptr
	Get_CurrentWindowVisualState      uintptr
	Get_CurrentWindowInteractionState uintptr
	Get_CachedCanMaximize             uintptr
	Get_CachedCanMinimize             uintptr
	Get_CachedIsModal                 uintptr
	Get_CachedIsTopmost               uintptr
	Get_CachedWindowVisualState       uintptr
	Get_CachedWindowInteractionState  uintptr
}

func (pat *UIAutomationWindowPattern) VTable() *UIAutomationWindowPatternVtbl {
	return (*UIAutomationWindowPatternVtbl)(unsafe.Pointer(pat.RawVTable))
}

func (pat *UIAutomationWindowPattern) Close() error {
	hr, _, _ := syscall.SyscallN(
		pat.VTable().Close,
		uintptr(unsafe.Pointer(pat)),
	)

	if hr != 0 {
		return ole.NewError(hr)
	}

	return nil
}

func (pat *UIAutomationWindowPattern) WaitForInputIdle(milliseconds int32) (bool, error) {
	var success bool

	hr, _, _ := syscall.SyscallN(
		pat.VTable().WaitForInputIdle,
		uintptr(unsafe.Pointer(pat)),
		uintptr(milliseconds),
		uintptr(unsafe.Pointer(&success)),
	)

	if hr != 0 {
		return false, ole.NewError(hr)
	}

	return success, nil
}

func (pat *UIAutomationWindowPattern) SetWindowVisualState(state WindowVisualState) error {
	hr, _, _ := syscall.SyscallN(
		pat.VTable().SetWindowVisualState,
		uintptr(unsafe.Pointer(pat)),
		uintptr(state),
	)

	if hr != 0 {
		return ole.NewError(hr)
	}

	return nil
}
