package uiautomation

import (
	"syscall"
	"unsafe"

	"github.com/go-ole/go-ole"
)

type PatternId uintptr

const (
	InvokePatternId            PatternId = 10000
	SelectionPatternId         PatternId = 10001
	ValuePatternId             PatternId = 10002
	RangeValuePatternId        PatternId = 10003
	ScrollPatternId            PatternId = 10004
	ExpandCollapsePatternId    PatternId = 10005
	GridPatternId              PatternId = 10006
	GridItemPatternId          PatternId = 10007
	MultipleViewPatternId      PatternId = 10008
	WindowPatternId            PatternId = 10009
	SelectionItemPatternId     PatternId = 10010
	DockPatternId              PatternId = 10011
	TablePatternId             PatternId = 10012
	TableItemPatternId         PatternId = 10013
	TextPatternId              PatternId = 10014
	TogglePatternId            PatternId = 10015
	TransformPatternId         PatternId = 10016
	ScrollItemPatternId        PatternId = 10017
	LegacyIAccessiblePatternId PatternId = 10018
	ItemContainerPatternId     PatternId = 10019
	VirtualizedItemPatternId   PatternId = 10020
	SynchronizedInputPatternId PatternId = 10021
	ObjectModelPatternId       PatternId = 10022
	AnnotationPatternId        PatternId = 10023
	TextPattern2Id             PatternId = 10024
	StylesPatternId            PatternId = 10025
	SpreadsheetPatternId       PatternId = 10026
	SpreadsheetItemPatternId   PatternId = 10027
	TransformPattern2Id        PatternId = 10028
	TextChildPatternId         PatternId = 10029
	DragPatternId              PatternId = 10030
	DropTargetPatternId        PatternId = 10031
	TextEditPatternId          PatternId = 10032
)

type InvokePattern struct {
	ole.IUnknown
}

type InvokePatternVtbl struct {
	ole.IUnknownVtbl
	Invoke uintptr
}

func (pat *InvokePattern) VTable() *InvokePatternVtbl {
	return (*InvokePatternVtbl)(unsafe.Pointer(pat.RawVTable))
}

func (pat *InvokePattern) Invoke() error {
	hr, _, _ := syscall.SyscallN(
		pat.VTable().Invoke,
		uintptr(unsafe.Pointer(pat)),
	)

	if hr != 0 {
		return ole.NewError(hr)
	}

	return nil
}

type SelectionPattern struct {
	ole.IUnknown
}

type SelectionPatternVtbl struct {
	ole.IUnknownVtbl
	GetCurrentSelection            uintptr
	Get_CurrentCanSelectMultiple   uintptr
	Get_CurrentIsSelectionRequired uintptr
	GetCachedSelection             uintptr
	Get_CachedCanSelectMultiple    uintptr
	Get_CachedIsSelectionRequired  uintptr
}

func (pat *SelectionPattern) VTable() *SelectionPatternVtbl {
	return (*SelectionPatternVtbl)(unsafe.Pointer(pat.RawVTable))
}

func (pat *SelectionPattern) GetCurrentSelection() (*ElementArray, error) {
	var selection *ElementArray

	hr, _, _ := syscall.SyscallN(
		pat.VTable().GetCurrentSelection,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&selection)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return selection, nil
}

func (pat *SelectionPattern) CurrentCanSelectMultiple() (bool, error) {
	var canSelectMultiple bool

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CurrentCanSelectMultiple,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&canSelectMultiple)),
	)

	if hr != 0 {
		return false, ole.NewError(hr)
	}

	return canSelectMultiple, nil
}

func (pat *SelectionPattern) CurrentIsSelectionRequired() (bool, error) {
	var isSelectionRequired bool

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CurrentIsSelectionRequired,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&isSelectionRequired)),
	)

	if hr != 0 {
		return false, ole.NewError(hr)
	}

	return isSelectionRequired, nil
}

func (pat *SelectionPattern) GetCachedSelection() (*ElementArray, error) {
	var selection *ElementArray

	hr, _, _ := syscall.SyscallN(
		pat.VTable().GetCachedSelection,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&selection)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return selection, nil
}

func (pat *SelectionPattern) CachedCanSelectMultiple() (bool, error) {
	var canSelectMultiple bool

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CachedCanSelectMultiple,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&canSelectMultiple)),
	)

	if hr != 0 {
		return false, ole.NewError(hr)
	}

	return canSelectMultiple, nil
}

func (pat *SelectionPattern) CachedIsSelectionRequired() (bool, error) {
	var isSelectionRequired bool

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CachedIsSelectionRequired,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&isSelectionRequired)),
	)

	if hr != 0 {
		return false, ole.NewError(hr)
	}

	return isSelectionRequired, nil
}

type ValuePattern struct {
	ole.IUnknown
}

type ValuePatternVtbl struct {
	ole.IUnknownVtbl
	SetValue                   uintptr
	Get_CurrentValue           uintptr
	Get_CurrentIsReadonly      uintptr
	Get_CachedValue            uintptr
	Get_CachedCachedIsReadOnly uintptr
}

func (pat *ValuePattern) VTable() *ValuePatternVtbl {
	return (*ValuePatternVtbl)(unsafe.Pointer(pat.RawVTable))
}

func (pat *ValuePattern) SetValue(value string) error {
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

func (pat *ValuePattern) CurrentValue() (string, error) {
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

func (pat *ValuePattern) CurrentIsReadonly() (bool, error) {
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

func (pat *ValuePattern) CachedValue() (string, error) {
	var value *uint16

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CachedValue,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&value)),
	)

	if hr != 0 {
		return "", ole.NewError(hr)
	}

	return ole.BstrToString(value), nil
}

func (pat *ValuePattern) CachedIsReadonly() (bool, error) {
	var readonly bool

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CachedCachedIsReadOnly,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&readonly)),
	)

	if hr != 0 {
		return false, ole.NewError(hr)
	}

	return readonly, nil
}

type RangeValuePattern struct {
	ole.IUnknown
}

type RangeValuePatternVtbl struct {
	ole.IUnknownVtbl
	SetValue               uintptr
	Get_CurrentValue       uintptr
	Get_CurrentIsReadOnly  uintptr
	Get_CurrentMaximum     uintptr
	Get_CurrentMinimum     uintptr
	Get_CurrentLargeChange uintptr
	Get_CurrentSmallChange uintptr
	Get_CachedValue        uintptr
	Get_CachedIsReadOnly   uintptr
	Get_CachedMaximum      uintptr
	Get_CachedMinimum      uintptr
	Get_CachedLargeChange  uintptr
	Get_CachedSmallChange  uintptr
}

func (pat *RangeValuePattern) VTable() *RangeValuePatternVtbl {
	return (*RangeValuePatternVtbl)(unsafe.Pointer(pat.RawVTable))
}

func (pat *RangeValuePattern) SetValue(value float64) error {
	hr, _, _ := syscall.SyscallN(
		pat.VTable().SetValue,
		uintptr(unsafe.Pointer(pat)),
		uintptr(value),
	)

	if hr != 0 {
		return ole.NewError(hr)
	}

	return nil
}

func (pat *RangeValuePattern) CurrentValue() (float64, error) {
	var value float64

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CurrentValue,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&value)),
	)

	if hr != 0 {
		return 0, ole.NewError(hr)
	}

	return value, nil
}

func (pat *RangeValuePattern) CurrentIsReadOnly() (bool, error) {
	var readonly bool

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CurrentIsReadOnly,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&readonly)),
	)

	if hr != 0 {
		return false, ole.NewError(hr)
	}

	return readonly, nil
}

func (pat *RangeValuePattern) CurrentMaximum() (float64, error) {
	var max float64

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CurrentMaximum,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&max)),
	)

	if hr != 0 {
		return 0, ole.NewError(hr)
	}

	return max, nil
}

func (pat *RangeValuePattern) CurrentMinimum() (float64, error) {
	var min float64

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CurrentMinimum,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&min)),
	)

	if hr != 0 {
		return 0, ole.NewError(hr)
	}

	return min, nil
}

func (pat *RangeValuePattern) CurrentLargeChange() (float64, error) {
	var largeChange float64

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CurrentLargeChange,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&largeChange)),
	)

	if hr != 0 {
		return 0, ole.NewError(hr)
	}

	return largeChange, nil
}

func (pat *RangeValuePattern) CurrentSmallChange() (float64, error) {
	var smallChange float64

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CurrentSmallChange,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&smallChange)),
	)

	if hr != 0 {
		return 0, ole.NewError(hr)
	}

	return smallChange, nil
}

func (pat *RangeValuePattern) CachedValue() (float64, error) {
	var value float64

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CachedValue,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&value)),
	)

	if hr != 0 {
		return 0, ole.NewError(hr)
	}

	return value, nil
}

func (pat *RangeValuePattern) CachedIsReadOnly() (bool, error) {
	var readonly bool

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CachedIsReadOnly,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&readonly)),
	)

	if hr != 0 {
		return false, ole.NewError(hr)
	}

	return readonly, nil
}

func (pat *RangeValuePattern) CachedMaximum() (float64, error) {
	var max float64

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CachedMaximum,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&max)),
	)

	if hr != 0 {
		return 0, ole.NewError(hr)
	}

	return max, nil
}

func (pat *RangeValuePattern) CachedMinimum() (float64, error) {
	var min float64

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CachedMinimum,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&min)),
	)

	if hr != 0 {
		return 0, ole.NewError(hr)
	}

	return min, nil
}

func (pat *RangeValuePattern) CachedLargeChange() (float64, error) {
	var largeChange float64

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CachedLargeChange,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&largeChange)),
	)

	if hr != 0 {
		return 0, ole.NewError(hr)
	}

	return largeChange, nil
}

func (pat *RangeValuePattern) CachedSmallChange() (float64, error) {
	var smallChange float64

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CachedSmallChange,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&smallChange)),
	)

	if hr != 0 {
		return 0, ole.NewError(hr)
	}

	return smallChange, nil
}

type WindowPattern struct {
	ole.IUnknown
}

type WindowPatternVtbl struct {
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

func (pat *WindowPattern) VTable() *WindowPatternVtbl {
	return (*WindowPatternVtbl)(unsafe.Pointer(pat.RawVTable))
}

func (pat *WindowPattern) Close() error {
	hr, _, _ := syscall.SyscallN(
		pat.VTable().Close,
		uintptr(unsafe.Pointer(pat)),
	)

	if hr != 0 {
		return ole.NewError(hr)
	}

	return nil
}

func (pat *WindowPattern) WaitForInputIdle(milliseconds int32) (bool, error) {
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

func (pat *WindowPattern) SetWindowVisualState(state WindowVisualState) error {
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

func (pat *WindowPattern) CurrentCanMaximize() (bool, error) {
	var canMaximize bool

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CurrentCanMaximize,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&canMaximize)),
	)

	if hr != 0 {
		return false, ole.NewError(hr)
	}

	return canMaximize, nil
}

func (pat *WindowPattern) CurrentCanMinimize() (bool, error) {
	var canMinimize bool

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CurrentCanMinimize,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&canMinimize)),
	)

	if hr != 0 {
		return false, ole.NewError(hr)
	}

	return canMinimize, nil
}

func (pat *WindowPattern) CurrentIsModal() (bool, error) {
	var isModal bool

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CurrentIsModal,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&isModal)),
	)

	if hr != 0 {
		return false, ole.NewError(hr)
	}

	return isModal, nil
}

func (pat *WindowPattern) CurrentIsTopmost() (bool, error) {
	var isTopmost bool

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CurrentIsTopmost,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&isTopmost)),
	)

	if hr != 0 {
		return false, ole.NewError(hr)
	}

	return isTopmost, nil
}

func (pat *WindowPattern) CurrentWindowVisualState() (WindowVisualState, error) {
	var state WindowVisualState

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CurrentWindowVisualState,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&state)),
	)

	if hr != 0 {
		return 0, ole.NewError(hr)
	}

	return state, nil
}

func (pat *WindowPattern) CurrentWindowInteractionState() (WindowInteractionState, error) {
	var state WindowInteractionState

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CurrentWindowInteractionState,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&state)),
	)

	if hr != 0 {
		return 0, ole.NewError(hr)
	}

	return state, nil
}

func (pat *WindowPattern) CachedCanMaximize() (bool, error) {
	var canMaximize bool

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CachedCanMaximize,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&canMaximize)),
	)

	if hr != 0 {
		return false, ole.NewError(hr)
	}

	return canMaximize, nil
}

func (pat *WindowPattern) CachedCanMinimize() (bool, error) {
	var canMinimize bool

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CachedCanMinimize,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&canMinimize)),
	)

	if hr != 0 {
		return false, ole.NewError(hr)
	}

	return canMinimize, nil
}

func (pat *WindowPattern) CachedIsModal() (bool, error) {
	var isModal bool

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CachedIsModal,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&isModal)),
	)

	if hr != 0 {
		return false, ole.NewError(hr)
	}

	return isModal, nil
}

func (pat *WindowPattern) CachedIsTopmost() (bool, error) {
	var isTopmost bool

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CachedIsTopmost,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&isTopmost)),
	)

	if hr != 0 {
		return false, ole.NewError(hr)
	}

	return isTopmost, nil
}

func (pat *WindowPattern) CachedWindowVisualState() (WindowVisualState, error) {
	var state WindowVisualState

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CachedWindowVisualState,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&state)),
	)

	if hr != 0 {
		return 0, ole.NewError(hr)
	}

	return state, nil
}

func (pat *WindowPattern) CachedWindowInteractionState() (WindowInteractionState, error) {
	var state WindowInteractionState

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CachedWindowInteractionState,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&state)),
	)

	if hr != 0 {
		return 0, ole.NewError(hr)
	}

	return state, nil
}

type TextPattern struct {
	ole.IUnknown
}

type TextPatternVtbl struct {
	ole.IUnknownVtbl
	RangeFromPoint             uintptr
	RangeFromChild             uintptr
	GetSelection               uintptr
	GetVisibleRanges           uintptr
	Get_DocumentRange          uintptr
	Get_SupportedTextSelection uintptr
}

func (pat *TextPattern) VTable() *TextPatternVtbl {
	return (*TextPatternVtbl)(unsafe.Pointer(pat.RawVTable))
}

func (pat *TextPattern) GetSelection() (*TextRangeArray, error) {
	var retVal *TextRangeArray

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

type ScrollItemPattern struct {
	ole.IUnknown
}

type ScrollItemPatternVtbl struct {
	ole.IUnknownVtbl
	ScrollIntoView uintptr
}

func (pat *ScrollItemPattern) VTable() *ScrollItemPatternVtbl {
	return (*ScrollItemPatternVtbl)(unsafe.Pointer(pat.RawVTable))
}

func (pat *ScrollItemPattern) ScrollIntoView() error {
	hr, _, _ := syscall.SyscallN(
		pat.VTable().ScrollIntoView,
		uintptr(unsafe.Pointer(pat)),
	)

	if hr != 0 {
		return ole.NewError(hr)
	}

	return nil
}

type LegacyAccessiblePattern struct {
	ole.IUnknown
}

type LegacyAccessiblePatternVtbl struct {
	ole.IUnknownVtbl
	Select                      uintptr
	DoDefaultAction             uintptr
	SetValue                    uintptr
	Get_CurrentChildId          uintptr
	Get_CurrentName             uintptr
	Get_CurrentValue            uintptr
	Get_CurrentDescription      uintptr
	Get_CurrentRole             uintptr
	Get_CurrentState            uintptr
	Get_CurrentHelp             uintptr
	Get_CurrentKeyboardShortcut uintptr
	GetCurrentSelection         uintptr
	Get_CurrentDefaultAction    uintptr
	Get_CachedChildId           uintptr
	Get_CachedName              uintptr
	Get_CachedValue             uintptr
	Get_CachedDescription       uintptr
	Get_CachedRole              uintptr
	Get_CachedState             uintptr
	Get_CachedHelp              uintptr
	Get_CachedKeyboardShortcut  uintptr
	GetCachedSelection          uintptr
	Get_CachedDefaultAction     uintptr
	GetIAccessible              uintptr
}

func (pat *LegacyAccessiblePattern) VTable() *LegacyAccessiblePatternVtbl {
	return (*LegacyAccessiblePatternVtbl)(unsafe.Pointer(pat.RawVTable))
}

func (pat *LegacyAccessiblePattern) Select(flags int32) error {
	hr, _, _ := syscall.SyscallN(
		pat.VTable().Select,
		uintptr(unsafe.Pointer(pat)),
		uintptr(flags),
	)

	if hr != 0 {
		return ole.NewError(hr)
	}

	return nil
}

func (pat *LegacyAccessiblePattern) DoDefaultAction() error {
	hr, _, _ := syscall.SyscallN(
		pat.VTable().DoDefaultAction,
		uintptr(unsafe.Pointer(pat)),
	)

	if hr != 0 {
		return ole.NewError(hr)
	}

	return nil
}

func (pat *LegacyAccessiblePattern) SetValue(value string) error {
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

func (pat *LegacyAccessiblePattern) CurrentChildId() (int32, error) {
	var childId int32

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CurrentChildId,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&childId)),
	)

	if hr != 0 {
		return 0, ole.NewError(hr)
	}

	return childId, nil
}

func (pat *LegacyAccessiblePattern) CurrentName() (string, error) {
	var name *uint16

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CurrentName,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&name)),
	)

	if hr != 0 {
		return "", ole.NewError(hr)
	}

	return ole.BstrToString(name), nil
}

func (pat *LegacyAccessiblePattern) CurrentValue() (string, error) {
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

func (pat *LegacyAccessiblePattern) CurrentDescription() (string, error) {
	var description *uint16

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CurrentDescription,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&description)),
	)

	if hr != 0 {
		return "", ole.NewError(hr)
	}

	return ole.BstrToString(description), nil
}

func (pat *LegacyAccessiblePattern) CurrentRole() (int32, error) {
	var role int32

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CurrentRole,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&role)),
	)

	if hr != 0 {
		return 0, ole.NewError(hr)
	}

	return role, nil
}

func (pat *LegacyAccessiblePattern) CurrentState() (int32, error) {
	var state int32

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CurrentState,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&state)),
	)

	if hr != 0 {
		return 0, ole.NewError(hr)
	}

	return state, nil
}

func (pat *LegacyAccessiblePattern) CurrentHelp() (string, error) {
	var help *uint16

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CurrentHelp,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&help)),
	)

	if hr != 0 {
		return "", ole.NewError(hr)
	}

	return ole.BstrToString(help), nil
}

func (pat *LegacyAccessiblePattern) CurrentKeyboardShortcut() (string, error) {
	var shortcut *uint16

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CurrentKeyboardShortcut,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&shortcut)),
	)

	if hr != 0 {
		return "", ole.NewError(hr)
	}

	return ole.BstrToString(shortcut), nil
}

func (pat *LegacyAccessiblePattern) GetCurrentSelection() (*ElementArray, error) {
	var selection *ElementArray

	hr, _, _ := syscall.SyscallN(
		pat.VTable().GetCurrentSelection,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&selection)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return selection, nil
}

func (pat *LegacyAccessiblePattern) CurrentDefaultAction() (string, error) {
	var action *uint16

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CurrentDefaultAction,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&action)),
	)

	if hr != 0 {
		return "", ole.NewError(hr)
	}

	return ole.BstrToString(action), nil
}

func (pat *LegacyAccessiblePattern) CachedChildId() (int32, error) {
	var childId int32

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CachedChildId,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&childId)),
	)

	if hr != 0 {
		return 0, ole.NewError(hr)
	}

	return childId, nil
}

func (pat *LegacyAccessiblePattern) CachedName() (string, error) {
	var name *uint16

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CachedName,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&name)),
	)

	if hr != 0 {
		return "", ole.NewError(hr)
	}

	return ole.BstrToString(name), nil
}

func (pat *LegacyAccessiblePattern) CachedValue() (string, error) {
	var value *uint16

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CachedValue,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&value)),
	)

	if hr != 0 {
		return "", ole.NewError(hr)
	}

	return ole.BstrToString(value), nil
}

func (pat *LegacyAccessiblePattern) CachedDescription() (string, error) {
	var description *uint16

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CachedDescription,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&description)),
	)

	if hr != 0 {
		return "", ole.NewError(hr)
	}

	return ole.BstrToString(description), nil
}

func (pat *LegacyAccessiblePattern) CachedRole() (int32, error) {
	var role int32

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CachedRole,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&role)),
	)

	if hr != 0 {
		return 0, ole.NewError(hr)
	}

	return role, nil
}

func (pat *LegacyAccessiblePattern) CachedState() (int32, error) {
	var state int32

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CachedState,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&state)),
	)

	if hr != 0 {
		return 0, ole.NewError(hr)
	}

	return state, nil
}

func (pat *LegacyAccessiblePattern) CachedHelp() (string, error) {
	var help *uint16

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CachedHelp,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&help)),
	)

	if hr != 0 {
		return "", ole.NewError(hr)
	}

	return ole.BstrToString(help), nil
}

func (pat *LegacyAccessiblePattern) CachedKeyboardShortcut() (string, error) {
	var shortcut *uint16

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CachedKeyboardShortcut,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&shortcut)),
	)

	if hr != 0 {
		return "", ole.NewError(hr)
	}

	return ole.BstrToString(shortcut), nil
}

func (pat *LegacyAccessiblePattern) GetCachedSelection() (*ElementArray, error) {
	var selection *ElementArray

	hr, _, _ := syscall.SyscallN(
		pat.VTable().GetCachedSelection,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&selection)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return selection, nil
}

func (pat *LegacyAccessiblePattern) CachedDefaultAction() (string, error) {
	var action *uint16

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CachedDefaultAction,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&action)),
	)

	if hr != 0 {
		return "", ole.NewError(hr)
	}

	return ole.BstrToString(action), nil
}

func (pat *LegacyAccessiblePattern) GetAccessible() (*Accessible, error) {
	var acc *Accessible

	hr, _, _ := syscall.SyscallN(
		pat.VTable().GetIAccessible,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&acc)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return acc, nil
}
