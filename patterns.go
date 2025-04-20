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
