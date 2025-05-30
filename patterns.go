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
	CustomNavigationPatternId  PatternId = 10033
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
	bstr := ole.SysAllocString(value)
	defer ole.SysFreeString(bstr)

	hr, _, _ := syscall.SyscallN(
		pat.VTable().SetValue,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(bstr)),
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

type ScrollPattern struct {
	ole.IUnknown
}

type ScrollPatternVtbl struct {
	ole.IUnknownVtbl
	Scroll                             uintptr
	SetScrollPercent                   uintptr
	Get_CurrentHorizontalScrollPercent uintptr
	Get_CurrentVerticalScrollPercent   uintptr
	Get_CurrentHorizontalViewSize      uintptr
	Get_CurrentVerticalViewSize        uintptr
	Get_CurrentHorizontallyScrollable  uintptr
	Get_CurrentVerticallyScrollable    uintptr
	Get_CachedHorizontalScrollPercent  uintptr
	Get_CachedVerticalScrollPercent    uintptr
	Get_CachedHorizontalViewSize       uintptr
	Get_CachedVerticalViewSize         uintptr
	Get_CachedHorizontallyScrollable   uintptr
	Get_CachedVerticallyScrollable     uintptr
}

func (pat *ScrollPattern) VTable() *ScrollPatternVtbl {
	return (*ScrollPatternVtbl)(unsafe.Pointer(pat.RawVTable))
}

func (pat *ScrollPattern) Scroll(horizontalAmount, verticalAmount ScrollAmount) error {
	hr, _, _ := syscall.SyscallN(
		pat.VTable().Scroll,
		uintptr(unsafe.Pointer(pat)),
		uintptr(horizontalAmount),
		uintptr(verticalAmount),
	)

	if hr != 0 {
		return ole.NewError(hr)
	}

	return nil
}

func (pat *ScrollPattern) SetScrollPercent(horizontalPercent, verticalPercent float64) error {
	hr, _, _ := syscall.SyscallN(
		pat.VTable().SetScrollPercent,
		uintptr(unsafe.Pointer(pat)),
		uintptr(horizontalPercent),
		uintptr(verticalPercent),
	)

	if hr != 0 {
		return ole.NewError(hr)
	}

	return nil
}

func (pat *ScrollPattern) CurrentHorizontalScrollPercent() (float64, error) {
	var percent float64

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CurrentHorizontalScrollPercent,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&percent)),
	)

	if hr != 0 {
		return 0, ole.NewError(hr)
	}

	return percent, nil
}

func (pat *ScrollPattern) CurrentVerticalScrollPercent() (float64, error) {
	var percent float64

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CurrentVerticalScrollPercent,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&percent)),
	)

	if hr != 0 {
		return 0, ole.NewError(hr)
	}

	return percent, nil
}

func (pat *ScrollPattern) CurrentHorizontalViewSize() (float64, error) {
	var size float64

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CurrentHorizontalViewSize,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&size)),
	)

	if hr != 0 {
		return 0, ole.NewError(hr)
	}

	return size, nil
}

func (pat *ScrollPattern) CurrentVerticalViewSize() (float64, error) {
	var size float64

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CurrentVerticalViewSize,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&size)),
	)

	if hr != 0 {
		return 0, ole.NewError(hr)
	}

	return size, nil
}

func (pat *ScrollPattern) CurrentHorizontallyScrollable() (bool, error) {
	var scrollable bool

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CurrentHorizontallyScrollable,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&scrollable)),
	)

	if hr != 0 {
		return false, ole.NewError(hr)
	}

	return scrollable, nil
}

func (pat *ScrollPattern) CurrentVerticallyScrollable() (bool, error) {
	var scrollable bool

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CurrentVerticallyScrollable,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&scrollable)),
	)

	if hr != 0 {
		return false, ole.NewError(hr)
	}

	return scrollable, nil
}

func (pat *ScrollPattern) CachedHorizontalScrollPercent() (float64, error) {
	var percent float64

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CachedHorizontalScrollPercent,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&percent)),
	)

	if hr != 0 {
		return 0, ole.NewError(hr)
	}

	return percent, nil
}

func (pat *ScrollPattern) CachedVerticalScrollPercent() (float64, error) {
	var percent float64

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CachedVerticalScrollPercent,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&percent)),
	)

	if hr != 0 {
		return 0, ole.NewError(hr)
	}

	return percent, nil
}

func (pat *ScrollPattern) CachedHorizontalViewSize() (float64, error) {
	var size float64

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CachedHorizontalViewSize,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&size)),
	)

	if hr != 0 {
		return 0, ole.NewError(hr)
	}

	return size, nil
}

func (pat *ScrollPattern) CachedVerticalViewSize() (float64, error) {
	var size float64

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CachedVerticalViewSize,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&size)),
	)

	if hr != 0 {
		return 0, ole.NewError(hr)
	}

	return size, nil
}

func (pat *ScrollPattern) CachedHorizontallyScrollable() (bool, error) {
	var scrollable bool

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CachedHorizontallyScrollable,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&scrollable)),
	)

	if hr != 0 {
		return false, ole.NewError(hr)
	}

	return scrollable, nil
}

func (pat *ScrollPattern) CachedVerticallyScrollable() (bool, error) {
	var scrollable bool

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CachedVerticallyScrollable,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&scrollable)),
	)

	if hr != 0 {
		return false, ole.NewError(hr)
	}

	return scrollable, nil
}

type ExpandCollapsePattern struct {
	ole.IUnknown
}

type ExpandCollapsePatternVtbl struct {
	ole.IUnknownVtbl
	Expand                         uintptr
	Collapse                       uintptr
	Get_CurrentExpandCollapseState uintptr
	Get_CachedExpandCollapseState  uintptr
}

func (pat *ExpandCollapsePattern) VTable() *ExpandCollapsePatternVtbl {
	return (*ExpandCollapsePatternVtbl)(unsafe.Pointer(pat.RawVTable))
}

func (pat *ExpandCollapsePattern) Expand() error {
	hr, _, _ := syscall.SyscallN(
		pat.VTable().Expand,
		uintptr(unsafe.Pointer(pat)),
	)

	if hr != 0 {
		return ole.NewError(hr)
	}

	return nil
}

func (pat *ExpandCollapsePattern) Collapse() error {
	hr, _, _ := syscall.SyscallN(
		pat.VTable().Collapse,
		uintptr(unsafe.Pointer(pat)),
	)

	if hr != 0 {
		return ole.NewError(hr)
	}

	return nil
}

func (pat *ExpandCollapsePattern) CurrentExpandCollapseState() (ExpandCollapseState, error) {
	var state ExpandCollapseState

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CurrentExpandCollapseState,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&state)),
	)

	if hr != 0 {
		return 0, ole.NewError(hr)
	}

	return state, nil
}

func (pat *ExpandCollapsePattern) CachedExpandCollapseState() (ExpandCollapseState, error) {
	var state ExpandCollapseState

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CachedExpandCollapseState,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&state)),
	)

	if hr != 0 {
		return 0, ole.NewError(hr)
	}

	return state, nil
}

type GridPattern struct {
	ole.IUnknown
}

type GridPatternVtbl struct {
	ole.IUnknownVtbl
	GetItem                uintptr
	Get_CurrentRowCount    uintptr
	Get_CurrentColumnCount uintptr
	Get_CachedRowCount     uintptr
	Get_CachedColumnCount  uintptr
}

func (pat *GridPattern) VTable() *GridPatternVtbl {
	return (*GridPatternVtbl)(unsafe.Pointer(pat.RawVTable))
}

func (pat *GridPattern) GetItem(row, column int32) (*Element, error) {
	var element *Element

	hr, _, _ := syscall.SyscallN(
		pat.VTable().GetItem,
		uintptr(unsafe.Pointer(pat)),
		uintptr(row),
		uintptr(column),
		uintptr(unsafe.Pointer(&element)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return element, nil
}

func (pat *GridPattern) CurrentRowCount() (int32, error) {
	var rowCount int32

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CurrentRowCount,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&rowCount)),
	)

	if hr != 0 {
		return 0, ole.NewError(hr)
	}

	return rowCount, nil
}

func (pat *GridPattern) CurrentColumnCount() (int32, error) {
	var columnCount int32

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CurrentColumnCount,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&columnCount)),
	)

	if hr != 0 {
		return 0, ole.NewError(hr)
	}

	return columnCount, nil
}

func (pat *GridPattern) CachedRowCount() (int32, error) {
	var rowCount int32

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CachedRowCount,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&rowCount)),
	)

	if hr != 0 {
		return 0, ole.NewError(hr)
	}

	return rowCount, nil
}

func (pat *GridPattern) CachedColumnCount() (int32, error) {
	var columnCount int32

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CachedColumnCount,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&columnCount)),
	)

	if hr != 0 {
		return 0, ole.NewError(hr)
	}

	return columnCount, nil
}

type GridItemPattern struct {
	ole.IUnknown
}

type GridItemPatternVtbl struct {
	ole.IUnknownVtbl
	Get_CurrentContainingGrid uintptr
	Get_CurrentRow            uintptr
	Get_CurrentColumn         uintptr
	Get_CurrentRowSpan        uintptr
	Get_CurrentColumnSpan     uintptr
	Get_CachedContainingGrid  uintptr
	Get_CachedRow             uintptr
	Get_CachedColumn          uintptr
	Get_CachedRowSpan         uintptr
	Get_CachedColumnSpan      uintptr
}

func (pat *GridItemPattern) VTable() *GridItemPatternVtbl {
	return (*GridItemPatternVtbl)(unsafe.Pointer(pat.RawVTable))
}

func (pat *GridItemPattern) CurrentContainingGrid() (*Element, error) {
	var grid *Element

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CurrentContainingGrid,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&grid)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return grid, nil
}

func (pat *GridItemPattern) CurrentRow() (int32, error) {
	var row int32

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CurrentRow,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&row)),
	)

	if hr != 0 {
		return 0, ole.NewError(hr)
	}

	return row, nil
}

func (pat *GridItemPattern) CurrentColumn() (int32, error) {
	var column int32

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CurrentColumn,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&column)),
	)

	if hr != 0 {
		return 0, ole.NewError(hr)
	}

	return column, nil
}

func (pat *GridItemPattern) CurrentRowSpan() (int32, error) {
	var rowSpan int32

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CurrentRowSpan,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&rowSpan)),
	)

	if hr != 0 {
		return 0, ole.NewError(hr)
	}

	return rowSpan, nil
}

func (pat *GridItemPattern) CurrentColumnSpan() (int32, error) {
	var columnSpan int32

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CurrentColumnSpan,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&columnSpan)),
	)

	if hr != 0 {
		return 0, ole.NewError(hr)
	}

	return columnSpan, nil
}

func (pat *GridItemPattern) CachedContainingGrid() (*Element, error) {
	var grid *Element

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CachedContainingGrid,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&grid)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return grid, nil
}

func (pat *GridItemPattern) CachedRow() (int32, error) {
	var row int32

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CachedRow,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&row)),
	)

	if hr != 0 {
		return 0, ole.NewError(hr)
	}

	return row, nil
}

func (pat *GridItemPattern) CachedColumn() (int32, error) {
	var column int32

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CachedColumn,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&column)),
	)

	if hr != 0 {
		return 0, ole.NewError(hr)
	}

	return column, nil
}

func (pat *GridItemPattern) CachedRowSpan() (int32, error) {
	var rowSpan int32

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CachedRowSpan,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&rowSpan)),
	)

	if hr != 0 {
		return 0, ole.NewError(hr)
	}

	return rowSpan, nil
}

func (pat *GridItemPattern) CachedColumnSpan() (int32, error) {
	var columnSpan int32

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CachedColumnSpan,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&columnSpan)),
	)

	if hr != 0 {
		return 0, ole.NewError(hr)
	}

	return columnSpan, nil
}

type MultipleViewPattern struct {
	ole.IUnknown
}

type MultipleViewPatternVtbl struct {
	ole.IUnknownVtbl
	GetViewName              uintptr
	SetCurrentView           uintptr
	Get_CurrentCurrentView   uintptr
	GetCurrentSupportedViews uintptr
	Get_CachedCurrentView    uintptr
	GetCachedSupportedViews  uintptr
}

func (pat *MultipleViewPattern) VTable() *MultipleViewPatternVtbl {
	return (*MultipleViewPatternVtbl)(unsafe.Pointer(pat.RawVTable))
}

func (pat *MultipleViewPattern) GetViewName(viewId int32) (string, error) {
	var name *uint16

	hr, _, _ := syscall.SyscallN(
		pat.VTable().GetViewName,
		uintptr(unsafe.Pointer(pat)),
		uintptr(viewId),
		uintptr(unsafe.Pointer(&name)),
	)

	if hr != 0 {
		return "", ole.NewError(hr)
	}

	return ole.BstrToString(name), nil
}

func (pat *MultipleViewPattern) SetCurrentView(viewId int32) error {
	hr, _, _ := syscall.SyscallN(
		pat.VTable().SetCurrentView,
		uintptr(unsafe.Pointer(pat)),
		uintptr(viewId),
	)

	if hr != 0 {
		return ole.NewError(hr)
	}

	return nil
}

func (pat *MultipleViewPattern) CurrentCurrentView() (int32, error) {
	var currentView int32

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CurrentCurrentView,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&currentView)),
	)

	if hr != 0 {
		return 0, ole.NewError(hr)
	}

	return currentView, nil
}

func (pat *MultipleViewPattern) CurrentSupportedViews() ([]int32, error) {
	var safeArray *ole.SafeArray

	hr, _, _ := syscall.SyscallN(
		pat.VTable().GetCurrentSupportedViews,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&safeArray)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	conversion := ole.SafeArrayConversion{Array: safeArray}
	defer conversion.Release()

	arr := conversion.ToValueArray()

	var views []int32
	for i := 0; i < len(views); i++ {
		if v, ok := arr[i].(int32); ok {
			views = append(views, v)
		}
	}

	return views, nil
}

func (pat *MultipleViewPattern) CachedCurrentView() (int32, error) {
	var currentView int32

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CachedCurrentView,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&currentView)),
	)

	if hr != 0 {
		return 0, ole.NewError(hr)
	}

	return currentView, nil
}

func (pat *MultipleViewPattern) CachedSupportedViews() ([]int32, error) {
	var safeArray *ole.SafeArray

	hr, _, _ := syscall.SyscallN(
		pat.VTable().GetCachedSupportedViews,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&safeArray)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	conversion := ole.SafeArrayConversion{Array: safeArray}
	defer conversion.Release()

	arr := conversion.ToValueArray()

	var views []int32
	for i := 0; i < len(views); i++ {
		if v, ok := arr[i].(int32); ok {
			views = append(views, v)
		}
	}

	return views, nil
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

type SelectionItemPattern struct {
	ole.IUnknown
}

type SelectionItemPatternVtbl struct {
	ole.IUnknownVtbl
	Select                        uintptr
	AddToSelection                uintptr
	RemoveFromSelection           uintptr
	Get_CurrentIsSelected         uintptr
	Get_CurrentSelectionContainer uintptr
	Get_CachedIsSelected          uintptr
	Get_CachedSelectionContainer  uintptr
}

func (pat *SelectionItemPattern) VTable() *SelectionItemPatternVtbl {
	return (*SelectionItemPatternVtbl)(unsafe.Pointer(pat.RawVTable))
}

func (pat *SelectionItemPattern) Select() error {
	hr, _, _ := syscall.SyscallN(
		pat.VTable().Select,
		uintptr(unsafe.Pointer(pat)),
	)

	if hr != 0 {
		return ole.NewError(hr)
	}

	return nil
}

func (pat *SelectionItemPattern) AddToSelection() error {
	hr, _, _ := syscall.SyscallN(
		pat.VTable().AddToSelection,
		uintptr(unsafe.Pointer(pat)),
	)

	if hr != 0 {
		return ole.NewError(hr)
	}

	return nil
}

func (pat *SelectionItemPattern) RemoveFromSelection() error {
	hr, _, _ := syscall.SyscallN(
		pat.VTable().RemoveFromSelection,
		uintptr(unsafe.Pointer(pat)),
	)

	if hr != 0 {
		return ole.NewError(hr)
	}

	return nil
}

func (pat *SelectionItemPattern) CurrentIsSelected() (bool, error) {
	var isSelected bool

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CurrentIsSelected,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&isSelected)),
	)

	if hr != 0 {
		return false, ole.NewError(hr)
	}

	return isSelected, nil
}

func (pat *SelectionItemPattern) CurrentSelectionContainer() (*Element, error) {
	var container *Element

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CurrentSelectionContainer,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&container)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return container, nil
}

func (pat *SelectionItemPattern) CachedIsSelected() (bool, error) {
	var isSelected bool

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CachedIsSelected,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&isSelected)),
	)

	if hr != 0 {
		return false, ole.NewError(hr)
	}

	return isSelected, nil
}

func (pat *SelectionItemPattern) CachedSelectionContainer() (*Element, error) {
	var container *Element

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CachedSelectionContainer,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&container)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return container, nil
}

type DockPattern struct {
	ole.IUnknown
}

type DockPatternVtbl struct {
	ole.IUnknownVtbl
	SetDockPosition         uintptr
	Get_CurrentDockPosition uintptr
	Get_CachedDockPosition  uintptr
}

func (pat *DockPattern) VTable() *DockPatternVtbl {
	return (*DockPatternVtbl)(unsafe.Pointer(pat.RawVTable))
}

func (pat *DockPattern) SetDockPosition(dockPosition DockPosition) error {
	hr, _, _ := syscall.SyscallN(
		pat.VTable().SetDockPosition,
		uintptr(unsafe.Pointer(pat)),
		uintptr(dockPosition),
	)

	if hr != 0 {
		return ole.NewError(hr)
	}

	return nil
}

func (pat *DockPattern) CurrentDockPosition() (DockPosition, error) {
	var dockPosition DockPosition

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CurrentDockPosition,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&dockPosition)),
	)

	if hr != 0 {
		return 0, ole.NewError(hr)
	}

	return dockPosition, nil
}

func (pat *DockPattern) CachedDockPosition() (DockPosition, error) {
	var dockPosition DockPosition

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CachedDockPosition,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&dockPosition)),
	)

	if hr != 0 {
		return 0, ole.NewError(hr)
	}

	return dockPosition, nil
}

type TablePattern struct {
	ole.IUnknown
}

type TablePatternVtbl struct {
	ole.IUnknownVtbl
	GetCurrentRowHeaders        uintptr
	GetCurrentColumnHeaders     uintptr
	Get_CurrentRowOrColumnMajor uintptr
	GetCachedRowHeaders         uintptr
	GetCachedColumnHeaders      uintptr
	Get_CachedRowOrColumnMajor  uintptr
}

func (pat *TablePattern) VTable() *TablePatternVtbl {
	return (*TablePatternVtbl)(unsafe.Pointer(pat.RawVTable))
}

func (pat *TablePattern) GetCurrentRowHeaders() (*ElementArray, error) {
	var headers *ElementArray

	hr, _, _ := syscall.SyscallN(
		pat.VTable().GetCurrentRowHeaders,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&headers)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return headers, nil
}

func (pat *TablePattern) GetCurrentColumnHeaders() (*ElementArray, error) {
	var headers *ElementArray

	hr, _, _ := syscall.SyscallN(
		pat.VTable().GetCurrentColumnHeaders,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&headers)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return headers, nil
}

func (pat *TablePattern) CurrentRowOrColumnMajor() (RowOrColumnMajor, error) {
	var major RowOrColumnMajor

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CurrentRowOrColumnMajor,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&major)),
	)

	if hr != 0 {
		return 0, ole.NewError(hr)
	}

	return major, nil
}

func (pat *TablePattern) GetCachedRowHeaders() (*ElementArray, error) {
	var headers *ElementArray

	hr, _, _ := syscall.SyscallN(
		pat.VTable().GetCachedRowHeaders,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&headers)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return headers, nil
}

func (pat *TablePattern) GetCachedColumnHeaders() (*ElementArray, error) {
	var headers *ElementArray

	hr, _, _ := syscall.SyscallN(
		pat.VTable().GetCachedColumnHeaders,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&headers)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return headers, nil
}

func (pat *TablePattern) CachedRowOrColumnMajor() (RowOrColumnMajor, error) {
	var major RowOrColumnMajor

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CachedRowOrColumnMajor,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&major)),
	)

	if hr != 0 {
		return 0, ole.NewError(hr)
	}

	return major, nil
}

type TableItemPattern struct {
	ole.IUnknown
}

type TableItemPatternVtbl struct {
	ole.IUnknownVtbl
	GetCurrentRowHeaderItems    uintptr
	GetCurrentColumnHeaderItems uintptr
	GetCachedRowHeaderItems     uintptr
	GetCachedColumnHeaderItems  uintptr
}

func (pat *TableItemPattern) VTable() *TableItemPatternVtbl {
	return (*TableItemPatternVtbl)(unsafe.Pointer(pat.RawVTable))
}

func (pat *TableItemPattern) GetCurrentRowHeaderItems() (*ElementArray, error) {
	var headers *ElementArray

	hr, _, _ := syscall.SyscallN(
		pat.VTable().GetCurrentRowHeaderItems,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&headers)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return headers, nil
}

func (pat *TableItemPattern) GetCurrentColumnHeaderItems() (*ElementArray, error) {
	var headers *ElementArray

	hr, _, _ := syscall.SyscallN(
		pat.VTable().GetCurrentColumnHeaderItems,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&headers)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return headers, nil
}

func (pat *TableItemPattern) GetCachedRowHeaderItems() (*ElementArray, error) {
	var headers *ElementArray

	hr, _, _ := syscall.SyscallN(
		pat.VTable().GetCachedRowHeaderItems,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&headers)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return headers, nil
}

func (pat *TableItemPattern) GetCachedColumnHeaderItems() (*ElementArray, error) {
	var headers *ElementArray

	hr, _, _ := syscall.SyscallN(
		pat.VTable().GetCachedColumnHeaderItems,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&headers)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return headers, nil
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

func (pat *TextPattern) RangeFromPoint(point Point) (*TextRange, error) {
	var rangeObj *TextRange

	hr, _, _ := syscall.SyscallN(
		pat.VTable().RangeFromPoint,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&point)),
		uintptr(unsafe.Pointer(&rangeObj)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return rangeObj, nil
}

func (pat *TextPattern) RangeFromChild(element *Element) (*TextRange, error) {
	var rangeObj *TextRange

	hr, _, _ := syscall.SyscallN(
		pat.VTable().RangeFromChild,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(element)),
		uintptr(unsafe.Pointer(&rangeObj)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return rangeObj, nil
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

func (pat *TextPattern) GetVisibleRanges() (*TextRangeArray, error) {
	var retVal *TextRangeArray

	hr, _, _ := syscall.SyscallN(
		pat.VTable().GetVisibleRanges,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&retVal)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return retVal, nil
}

func (pat *TextPattern) DocumentRange() (*TextRange, error) {
	var rangeObj *TextRange

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_DocumentRange,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&rangeObj)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return rangeObj, nil
}

func (pat *TextPattern) SupportedTextSelection() (SupportedTextSelection, error) {
	var selection SupportedTextSelection

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_SupportedTextSelection,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&selection)),
	)

	if hr != 0 {
		return 0, ole.NewError(hr)
	}

	return selection, nil
}

type TogglePattern struct {
	ole.IUnknown
}

type TogglePatternVtbl struct {
	ole.IUnknownVtbl
	Toggle                 uintptr
	Get_CurrentToggleState uintptr
	Get_CachedToggleState  uintptr
}

func (pat *TogglePattern) VTable() *TogglePatternVtbl {
	return (*TogglePatternVtbl)(unsafe.Pointer(pat.RawVTable))
}

func (pat *TogglePattern) Toggle() error {
	hr, _, _ := syscall.SyscallN(
		pat.VTable().Toggle,
		uintptr(unsafe.Pointer(pat)),
	)

	if hr != 0 {
		return ole.NewError(hr)
	}

	return nil
}

func (pat *TogglePattern) CurrentToggleState() (ToggleState, error) {
	var state ToggleState

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CurrentToggleState,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&state)),
	)

	if hr != 0 {
		return 0, ole.NewError(hr)
	}

	return state, nil
}

func (pat *TogglePattern) CachedToggleState() (ToggleState, error) {
	var state ToggleState

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CachedToggleState,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&state)),
	)

	if hr != 0 {
		return 0, ole.NewError(hr)
	}

	return state, nil
}

type TransformPattern struct {
	ole.IUnknown
}

type TransformPatternVtbl struct {
	ole.IUnknownVtbl
	Move                 uintptr
	Resize               uintptr
	Rotate               uintptr
	Get_CurrentCanMove   uintptr
	Get_CurrentCanResize uintptr
	Get_CurrentCanRotate uintptr
	Get_CachedCanMove    uintptr
	Get_CachedCanResize  uintptr
	Get_CachedCanRotate  uintptr
}

func (pat *TransformPattern) VTable() *TransformPatternVtbl {
	return (*TransformPatternVtbl)(unsafe.Pointer(pat.RawVTable))
}

func (pat *TransformPattern) Move(x, y float64) error {
	hr, _, _ := syscall.SyscallN(
		pat.VTable().Move,
		uintptr(unsafe.Pointer(pat)),
		uintptr(x),
		uintptr(y),
	)

	if hr != 0 {
		return ole.NewError(hr)
	}

	return nil
}

func (pat *TransformPattern) Resize(width, height float64) error {
	hr, _, _ := syscall.SyscallN(
		pat.VTable().Resize,
		uintptr(unsafe.Pointer(pat)),
		uintptr(width),
		uintptr(height),
	)

	if hr != 0 {
		return ole.NewError(hr)
	}

	return nil
}

func (pat *TransformPattern) Rotate(degrees float64) error {
	hr, _, _ := syscall.SyscallN(
		pat.VTable().Rotate,
		uintptr(unsafe.Pointer(pat)),
		uintptr(degrees),
	)

	if hr != 0 {
		return ole.NewError(hr)
	}

	return nil
}

func (pat *TransformPattern) CurrentCanMove() (bool, error) {
	var canMove bool

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CurrentCanMove,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&canMove)),
	)

	if hr != 0 {
		return false, ole.NewError(hr)
	}

	return canMove, nil
}

func (pat *TransformPattern) CurrentCanResize() (bool, error) {
	var canResize bool

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CurrentCanResize,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&canResize)),
	)

	if hr != 0 {
		return false, ole.NewError(hr)
	}

	return canResize, nil
}

func (pat *TransformPattern) CurrentCanRotate() (bool, error) {
	var canRotate bool

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CurrentCanRotate,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&canRotate)),
	)

	if hr != 0 {
		return false, ole.NewError(hr)
	}

	return canRotate, nil
}

func (pat *TransformPattern) CachedCanMove() (bool, error) {
	var canMove bool

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CachedCanMove,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&canMove)),
	)

	if hr != 0 {
		return false, ole.NewError(hr)
	}

	return canMove, nil
}

func (pat *TransformPattern) CachedCanResize() (bool, error) {
	var canResize bool

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CachedCanResize,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&canResize)),
	)

	if hr != 0 {
		return false, ole.NewError(hr)
	}

	return canResize, nil
}

func (pat *TransformPattern) CachedCanRotate() (bool, error) {
	var canRotate bool

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CachedCanRotate,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&canRotate)),
	)

	if hr != 0 {
		return false, ole.NewError(hr)
	}

	return canRotate, nil
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
	bstr := ole.SysAllocString(value)
	defer ole.SysFreeString(bstr)

	hr, _, _ := syscall.SyscallN(
		pat.VTable().SetValue,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(bstr)),
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

type ItemContainerPattern struct {
	ole.IUnknown
}

type ItemContainerPatternVtbl struct {
	ole.IUnknownVtbl
	FindItemByProperty uintptr
}

func (pat *ItemContainerPattern) VTable() *ItemContainerPatternVtbl {
	return (*ItemContainerPatternVtbl)(unsafe.Pointer(pat.RawVTable))
}

func (pat *ItemContainerPattern) FindItemByProperty(startAfter *Element, propertyId PropertyId, value ole.VARIANT) (*Element, error) {
	var item *Element

	hr, _, _ := syscall.SyscallN(
		pat.VTable().FindItemByProperty,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(startAfter)),
		uintptr(propertyId),
		uintptr(unsafe.Pointer(&value)),
		uintptr(unsafe.Pointer(&item)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return item, nil
}

type VirtualizedItemPattern struct {
	ole.IUnknown
}

type VirtualizedItemPatternVtbl struct {
	ole.IUnknownVtbl
	Realize uintptr
}

func (pat *VirtualizedItemPattern) VTable() *VirtualizedItemPatternVtbl {
	return (*VirtualizedItemPatternVtbl)(unsafe.Pointer(pat.RawVTable))
}

func (pat *VirtualizedItemPattern) Realize() error {
	hr, _, _ := syscall.SyscallN(
		pat.VTable().Realize,
		uintptr(unsafe.Pointer(pat)),
	)

	if hr != 0 {
		return ole.NewError(hr)
	}

	return nil
}

type SynchronizedInputPattern struct {
	ole.IUnknown
}

type SynchronizedInputPatternVtbl struct {
	ole.IUnknownVtbl
	StartListening uintptr
	Cancel         uintptr
}

func (pat *SynchronizedInputPattern) VTable() *SynchronizedInputPatternVtbl {
	return (*SynchronizedInputPatternVtbl)(unsafe.Pointer(pat.RawVTable))
}

func (pat *SynchronizedInputPattern) StartListening(inputType SynchronizedInputType) error {
	hr, _, _ := syscall.SyscallN(
		pat.VTable().StartListening,
		uintptr(unsafe.Pointer(pat)),
		uintptr(inputType),
	)

	if hr != 0 {
		return ole.NewError(hr)
	}

	return nil
}

func (pat *SynchronizedInputPattern) Cancel() error {
	hr, _, _ := syscall.SyscallN(
		pat.VTable().Cancel,
		uintptr(unsafe.Pointer(pat)),
	)

	if hr != 0 {
		return ole.NewError(hr)
	}

	return nil
}

type ObjectModelPattern struct {
	ole.IUnknown
}

type ObjectModelPatternVtbl struct {
	ole.IUnknownVtbl
	GetUnderlyingObjectModel uintptr
}

func (pat *ObjectModelPattern) VTable() *ObjectModelPatternVtbl {
	return (*ObjectModelPatternVtbl)(unsafe.Pointer(pat.RawVTable))
}

func (pat *ObjectModelPattern) GetUnderlyingObjectModel() (*ole.IUnknown, error) {
	var obj *ole.IUnknown

	hr, _, _ := syscall.SyscallN(
		pat.VTable().GetUnderlyingObjectModel,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&obj)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return obj, nil
}

type AnnotationPattern struct {
	ole.IUnknown
}

type AnnotationPatternVtbl struct {
	ole.IUnknownVtbl
	Get_CurrentAnnotationTypeId   uintptr
	Get_CurrentAnnotationTypeName uintptr
	Get_CurrentAuthor             uintptr
	Get_CurrentDateTime           uintptr
	Get_CurrentTarget             uintptr
	Get_CachedAnnotationTypeId    uintptr
	Get_CachedAnnotationTypeName  uintptr
	Get_CachedAuthor              uintptr
	Get_CachedDateTime            uintptr
	Get_CachedTarget              uintptr
}

func (pat *AnnotationPattern) VTable() *AnnotationPatternVtbl {
	return (*AnnotationPatternVtbl)(unsafe.Pointer(pat.RawVTable))
}

func (pat *AnnotationPattern) CurrentAnnotationTypeId() (int32, error) {
	var typeId int32

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CurrentAnnotationTypeId,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&typeId)),
	)

	if hr != 0 {
		return 0, ole.NewError(hr)
	}

	return typeId, nil
}

func (pat *AnnotationPattern) CurrentAnnotationTypeName() (string, error) {
	var typeName *uint16

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CurrentAnnotationTypeName,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&typeName)),
	)

	if hr != 0 {
		return "", ole.NewError(hr)
	}

	return ole.BstrToString(typeName), nil
}

func (pat *AnnotationPattern) CurrentAuthor() (string, error) {
	var author *uint16

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CurrentAuthor,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&author)),
	)

	if hr != 0 {
		return "", ole.NewError(hr)
	}

	return ole.BstrToString(author), nil
}

func (pat *AnnotationPattern) CurrentDateTime() (string, error) {
	var dateTime *uint16

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CurrentDateTime,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&dateTime)),
	)

	if hr != 0 {
		return "", ole.NewError(hr)
	}

	return ole.BstrToString(dateTime), nil
}

func (pat *AnnotationPattern) CurrentTarget() (*Element, error) {
	var target *Element

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CurrentTarget,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&target)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return target, nil
}

func (pat *AnnotationPattern) CachedAnnotationTypeId() (int32, error) {
	var typeId int32

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CachedAnnotationTypeId,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&typeId)),
	)

	if hr != 0 {
		return 0, ole.NewError(hr)
	}

	return typeId, nil
}

func (pat *AnnotationPattern) CachedAnnotationTypeName() (string, error) {
	var typeName *uint16

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CachedAnnotationTypeName,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&typeName)),
	)

	if hr != 0 {
		return "", ole.NewError(hr)
	}

	return ole.BstrToString(typeName), nil
}

func (pat *AnnotationPattern) CachedAuthor() (string, error) {
	var author *uint16

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CachedAuthor,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&author)),
	)

	if hr != 0 {
		return "", ole.NewError(hr)
	}

	return ole.BstrToString(author), nil
}

func (pat *AnnotationPattern) CachedDateTime() (string, error) {
	var dateTime *uint16

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CachedDateTime,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&dateTime)),
	)

	if hr != 0 {
		return "", ole.NewError(hr)
	}

	return ole.BstrToString(dateTime), nil
}

func (pat *AnnotationPattern) CachedTarget() (*Element, error) {
	var target *Element

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CachedTarget,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&target)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return target, nil
}

type TextPattern2 struct {
	TextPattern
}

type TextPattern2Vtbl struct {
	TextPatternVtbl
	RangeFromAnnotation uintptr
	GetCaretRange       uintptr
}

func (pat *TextPattern2) VTable() *TextPattern2Vtbl {
	return (*TextPattern2Vtbl)(unsafe.Pointer(pat.RawVTable))
}

func (pat *TextPattern2) RangeFromAnnotation(annotation *Element) (*TextRange, error) {
	var rangeObj *TextRange

	hr, _, _ := syscall.SyscallN(
		pat.VTable().RangeFromAnnotation,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(annotation)),
		uintptr(unsafe.Pointer(&rangeObj)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return rangeObj, nil
}

func (pat *TextPattern2) GetCaretRange() (bool, *TextRange, error) {
	var isActive bool
	var rangeObj *TextRange

	hr, _, _ := syscall.SyscallN(
		pat.VTable().GetCaretRange,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&isActive)),
		uintptr(unsafe.Pointer(&rangeObj)),
	)

	if hr != 0 {
		return false, nil, ole.NewError(hr)
	}

	return isActive, rangeObj, nil
}

type StylesPattern struct {
	ole.IUnknown
}

type StylesPatternVtbl struct {
	ole.IUnknownVtbl
	Get_CurrentStyleId                  uintptr
	Get_CurrentStyleName                uintptr
	Get_CurrentFillColor                uintptr
	Get_CurrentFillPatternStyle         uintptr
	Get_CurrentShape                    uintptr
	Get_CurrentFillPatternColor         uintptr
	Get_CurrentExtendedProperties       uintptr
	GetCurrentExtendedPropertiesAsArray uintptr
	Get_CachedStyleId                   uintptr
	Get_CachedStyleName                 uintptr
	Get_CachedFillColor                 uintptr
	Get_CachedFillPatternStyle          uintptr
	Get_CachedShape                     uintptr
	Get_CachedFillPatternColor          uintptr
	Get_CachedExtendedProperties        uintptr
	GetCachedExtendedPropertiesAsArray  uintptr
}

func (pat *StylesPattern) VTable() *StylesPatternVtbl {
	return (*StylesPatternVtbl)(unsafe.Pointer(pat.RawVTable))
}

func (pat *StylesPattern) CurrentStyleId() (int32, error) {
	var styleId int32

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CurrentStyleId,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&styleId)),
	)

	if hr != 0 {
		return 0, ole.NewError(hr)
	}

	return styleId, nil
}

func (pat *StylesPattern) CurrentStyleName() (string, error) {
	var styleName *uint16

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CurrentStyleName,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&styleName)),
	)

	if hr != 0 {
		return "", ole.NewError(hr)
	}

	return ole.BstrToString(styleName), nil
}

func (pat *StylesPattern) CurrentFillColor() (int32, error) {
	var fillColor int32

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CurrentFillColor,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&fillColor)),
	)

	if hr != 0 {
		return 0, ole.NewError(hr)
	}

	return fillColor, nil
}

func (pat *StylesPattern) CurrentFillPatternStyle() (string, error) {
	var patternStyle *uint16

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CurrentFillPatternStyle,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&patternStyle)),
	)

	if hr != 0 {
		return "", ole.NewError(hr)
	}

	return ole.BstrToString(patternStyle), nil
}

func (pat *StylesPattern) CurrentShape() (string, error) {
	var shape *uint16

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CurrentShape,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&shape)),
	)

	if hr != 0 {
		return "", ole.NewError(hr)
	}

	return ole.BstrToString(shape), nil
}

func (pat *StylesPattern) CurrentFillPatternColor() (int32, error) {
	var patternColor int32

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CurrentFillPatternColor,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&patternColor)),
	)

	if hr != 0 {
		return 0, ole.NewError(hr)
	}

	return patternColor, nil
}

func (pat *StylesPattern) CurrentExtendedProperties() (string, error) {
	var properties *uint16

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CurrentExtendedProperties,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&properties)),
	)

	if hr != 0 {
		return "", ole.NewError(hr)
	}

	return ole.BstrToString(properties), nil
}

func (pat *StylesPattern) GetCurrentExtendedPropertiesAsArray() ([]ExtendedProperty, error) {
	var propertiesPtr *ExtendedProperty
	var count int32

	hr, _, _ := syscall.SyscallN(
		pat.VTable().GetCurrentExtendedPropertiesAsArray,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&propertiesPtr)),
		uintptr(unsafe.Pointer(&count)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	slice := (*[1 << 30]ExtendedProperty)(unsafe.Pointer(propertiesPtr))[:count:count]
	properties := make([]ExtendedProperty, count)
	copy(properties, slice)

	ole.CoTaskMemFree(uintptr(unsafe.Pointer(propertiesPtr)))

	return properties, nil
}

func (pat *StylesPattern) CachedStyleId() (int32, error) {
	var styleId int32

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CachedStyleId,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&styleId)),
	)

	if hr != 0 {
		return 0, ole.NewError(hr)
	}

	return styleId, nil
}

func (pat *StylesPattern) CachedStyleName() (string, error) {
	var styleName *uint16

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CachedStyleName,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&styleName)),
	)

	if hr != 0 {
		return "", ole.NewError(hr)
	}

	return ole.BstrToString(styleName), nil
}

func (pat *StylesPattern) CachedFillColor() (int32, error) {
	var fillColor int32

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CachedFillColor,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&fillColor)),
	)

	if hr != 0 {
		return 0, ole.NewError(hr)
	}

	return fillColor, nil
}

func (pat *StylesPattern) CachedFillPatternStyle() (string, error) {
	var patternStyle *uint16

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CachedFillPatternStyle,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&patternStyle)),
	)

	if hr != 0 {
		return "", ole.NewError(hr)
	}

	return ole.BstrToString(patternStyle), nil
}

func (pat *StylesPattern) CachedShape() (string, error) {
	var shape *uint16

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CachedShape,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&shape)),
	)

	if hr != 0 {
		return "", ole.NewError(hr)
	}

	return ole.BstrToString(shape), nil
}

func (pat *StylesPattern) CachedFillPatternColor() (int32, error) {
	var patternColor int32

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CachedFillPatternColor,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&patternColor)),
	)

	if hr != 0 {
		return 0, ole.NewError(hr)
	}

	return patternColor, nil
}

func (pat *StylesPattern) CachedExtendedProperties() (string, error) {
	var properties *uint16

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CachedExtendedProperties,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&properties)),
	)

	if hr != 0 {
		return "", ole.NewError(hr)
	}

	return ole.BstrToString(properties), nil
}

func (pat *StylesPattern) GetCachedExtendedPropertiesAsArray() ([]ExtendedProperty, error) {
	var propertiesPtr *ExtendedProperty
	var count int32

	hr, _, _ := syscall.SyscallN(
		pat.VTable().GetCachedExtendedPropertiesAsArray,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&propertiesPtr)),
		uintptr(unsafe.Pointer(&count)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	slice := (*[1 << 30]ExtendedProperty)(unsafe.Pointer(propertiesPtr))[:count:count]
	properties := make([]ExtendedProperty, count)
	copy(properties, slice)

	ole.CoTaskMemFree(uintptr(unsafe.Pointer(propertiesPtr)))

	return properties, nil
}

type SpreadsheetPattern struct {
	ole.IUnknown
}

type SpreadsheetPatternVtbl struct {
	ole.IUnknownVtbl
	GetItemByName uintptr
}

func (pat *SpreadsheetPattern) VTable() *SpreadsheetPatternVtbl {
	return (*SpreadsheetPatternVtbl)(unsafe.Pointer(pat.RawVTable))
}

func (pat *SpreadsheetPattern) GetItemByName(name string) (*Element, error) {
	bstr := ole.SysAllocString(name)
	defer ole.SysFreeString(bstr)

	var item *Element

	hr, _, _ := syscall.SyscallN(
		pat.VTable().GetItemByName,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(bstr)),
		uintptr(unsafe.Pointer(&item)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return item, nil
}

type SpreadsheetItemPattern struct {
	ole.IUnknown
}

type SpreadsheetItemPatternVtbl struct {
	ole.IUnknownVtbl
	Get_CurrentFormula          uintptr
	GetCurrentAnnotationObjects uintptr
	GetCurrentAnnotationTypes   uintptr
	Get_CachedFormula           uintptr
	GetCachedAnnotationObjects  uintptr
	GetCachedAnnotationTypes    uintptr
}

func (pat *SpreadsheetItemPattern) VTable() *SpreadsheetItemPatternVtbl {
	return (*SpreadsheetItemPatternVtbl)(unsafe.Pointer(pat.RawVTable))
}

func (pat *SpreadsheetItemPattern) CurrentFormula() (string, error) {
	var formula *uint16

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CurrentFormula,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&formula)),
	)

	if hr != 0 {
		return "", ole.NewError(hr)
	}

	return ole.BstrToString(formula), nil
}

func (pat *SpreadsheetItemPattern) GetCurrentAnnotationObjects() (*ElementArray, error) {
	var objects *ElementArray

	hr, _, _ := syscall.SyscallN(
		pat.VTable().GetCurrentAnnotationObjects,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&objects)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return objects, nil
}

func (pat *SpreadsheetItemPattern) GetCurrentAnnotationTypes() (*ole.SafeArray, error) {
	var types *ole.SafeArray

	hr, _, _ := syscall.SyscallN(
		pat.VTable().GetCurrentAnnotationTypes,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&types)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return types, nil
}

func (pat *SpreadsheetItemPattern) CachedFormula() (string, error) {
	var formula *uint16

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CachedFormula,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&formula)),
	)

	if hr != 0 {
		return "", ole.NewError(hr)
	}

	return ole.BstrToString(formula), nil
}

func (pat *SpreadsheetItemPattern) GetCachedAnnotationObjects() (*ElementArray, error) {
	var objects *ElementArray

	hr, _, _ := syscall.SyscallN(
		pat.VTable().GetCachedAnnotationObjects,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&objects)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return objects, nil
}

func (pat *SpreadsheetItemPattern) GetCachedAnnotationTypes() (*ole.SafeArray, error) {
	var types *ole.SafeArray

	hr, _, _ := syscall.SyscallN(
		pat.VTable().GetCachedAnnotationTypes,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&types)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return types, nil
}

type TransformPattern2 struct {
	TransformPattern
}

type TransformPattern2Vtbl struct {
	TransformPatternVtbl
	Zoom                   uintptr
	ZoomByUnit             uintptr
	Get_CurrentCanZoom     uintptr
	Get_CachedCanZoom      uintptr
	Get_CurrentZoomLevel   uintptr
	Get_CachedZoomLevel    uintptr
	Get_CurrentZoomMinimum uintptr
	Get_CachedZoomMinimum  uintptr
	Get_CurrentZoomMaximum uintptr
	Get_CachedZoomMaximum  uintptr
}

func (pat *TransformPattern2) VTable() *TransformPattern2Vtbl {
	return (*TransformPattern2Vtbl)(unsafe.Pointer(pat.RawVTable))
}

func (pat *TransformPattern2) Zoom(zoomValue float64) error {
	hr, _, _ := syscall.SyscallN(
		pat.VTable().Zoom,
		uintptr(unsafe.Pointer(pat)),
		uintptr(zoomValue),
	)

	if hr != 0 {
		return ole.NewError(hr)
	}

	return nil
}

func (pat *TransformPattern2) ZoomByUnit(zoomUnit ZoomUnit) error {
	hr, _, _ := syscall.SyscallN(
		pat.VTable().ZoomByUnit,
		uintptr(unsafe.Pointer(pat)),
		uintptr(zoomUnit),
	)

	if hr != 0 {
		return ole.NewError(hr)
	}

	return nil
}

func (pat *TransformPattern2) CurrentCanZoom() (bool, error) {
	var canZoom bool

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CurrentCanZoom,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&canZoom)),
	)

	if hr != 0 {
		return false, ole.NewError(hr)
	}

	return canZoom, nil
}

func (pat *TransformPattern2) CachedCanZoom() (bool, error) {
	var canZoom bool

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CachedCanZoom,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&canZoom)),
	)

	if hr != 0 {
		return false, ole.NewError(hr)
	}

	return canZoom, nil
}

func (pat *TransformPattern2) CurrentZoomLevel() (float64, error) {
	var zoomLevel float64

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CurrentZoomLevel,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&zoomLevel)),
	)

	if hr != 0 {
		return 0, ole.NewError(hr)
	}

	return zoomLevel, nil
}

func (pat *TransformPattern2) CachedZoomLevel() (float64, error) {
	var zoomLevel float64

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CachedZoomLevel,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&zoomLevel)),
	)

	if hr != 0 {
		return 0, ole.NewError(hr)
	}

	return zoomLevel, nil
}

func (pat *TransformPattern2) CurrentZoomMinimum() (float64, error) {
	var zoomMin float64

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CurrentZoomMinimum,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&zoomMin)),
	)

	if hr != 0 {
		return 0, ole.NewError(hr)
	}

	return zoomMin, nil
}

func (pat *TransformPattern2) CachedZoomMinimum() (float64, error) {
	var zoomMin float64

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CachedZoomMinimum,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&zoomMin)),
	)

	if hr != 0 {
		return 0, ole.NewError(hr)
	}

	return zoomMin, nil
}

func (pat *TransformPattern2) CurrentZoomMaximum() (float64, error) {
	var zoomMax float64

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CurrentZoomMaximum,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&zoomMax)),
	)

	if hr != 0 {
		return 0, ole.NewError(hr)
	}

	return zoomMax, nil
}

func (pat *TransformPattern2) CachedZoomMaximum() (float64, error) {
	var zoomMax float64

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CachedZoomMaximum,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&zoomMax)),
	)

	if hr != 0 {
		return 0, ole.NewError(hr)
	}

	return zoomMax, nil
}

type TextChildPattern struct {
	ole.IUnknown
}

type TextChildPatternVtbl struct {
	ole.IUnknownVtbl
	Get_TextContainer uintptr
	Get_TextRange     uintptr
}

func (pat *TextChildPattern) VTable() *TextChildPatternVtbl {
	return (*TextChildPatternVtbl)(unsafe.Pointer(pat.RawVTable))
}

func (pat *TextChildPattern) GetTextContainer() (*Element, error) {
	var container *Element

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_TextContainer,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&container)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return container, nil
}

func (pat *TextChildPattern) GetTextRange() (*TextRange, error) {
	var rangeObj *TextRange

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_TextRange,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&rangeObj)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return rangeObj, nil
}

type DragPattern struct {
	ole.IUnknown
}

type DragPatternVtbl struct {
	ole.IUnknownVtbl
	Get_CurrentIsGrabbed   uintptr
	Get_CachedIsGrabbed    uintptr
	Get_CurrentDropEffect  uintptr
	Get_CachedDropEffect   uintptr
	Get_CurrentDropEffects uintptr
	Get_CachedDropEffects  uintptr
	GetCurrentGrabbedItems uintptr
	GetCachedGrabbedItems  uintptr
}

func (pat *DragPattern) VTable() *DragPatternVtbl {
	return (*DragPatternVtbl)(unsafe.Pointer(pat.RawVTable))
}

func (pat *DragPattern) CurrentIsGrabbed() (bool, error) {
	var isGrabbed bool

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CurrentIsGrabbed,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&isGrabbed)),
	)

	if hr != 0 {
		return false, ole.NewError(hr)
	}

	return isGrabbed, nil
}

func (pat *DragPattern) CachedIsGrabbed() (bool, error) {
	var isGrabbed bool

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CachedIsGrabbed,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&isGrabbed)),
	)

	if hr != 0 {
		return false, ole.NewError(hr)
	}

	return isGrabbed, nil
}

func (pat *DragPattern) CurrentDropEffect() (string, error) {
	var dropEffect *uint16

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CurrentDropEffect,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&dropEffect)),
	)

	if hr != 0 {
		return "", ole.NewError(hr)
	}

	return ole.BstrToString(dropEffect), nil
}

func (pat *DragPattern) CachedDropEffect() (string, error) {
	var dropEffect *uint16

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CachedDropEffect,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&dropEffect)),
	)

	if hr != 0 {
		return "", ole.NewError(hr)
	}

	return ole.BstrToString(dropEffect), nil
}

func (pat *DragPattern) CurrentDropEffects() (*ole.SafeArray, error) {
	var dropEffects *ole.SafeArray

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CurrentDropEffects,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&dropEffects)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return dropEffects, nil
}

func (pat *DragPattern) CachedDropEffects() (*ole.SafeArray, error) {
	var dropEffects *ole.SafeArray

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CachedDropEffects,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&dropEffects)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return dropEffects, nil
}

func (pat *DragPattern) GetCurrentGrabbedItems() (*ElementArray, error) {
	var grabbedItems *ElementArray

	hr, _, _ := syscall.SyscallN(
		pat.VTable().GetCurrentGrabbedItems,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&grabbedItems)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return grabbedItems, nil
}

func (pat *DragPattern) GetCachedGrabbedItems() (*ElementArray, error) {
	var grabbedItems *ElementArray

	hr, _, _ := syscall.SyscallN(
		pat.VTable().GetCachedGrabbedItems,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&grabbedItems)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return grabbedItems, nil
}

type DropTargetPattern struct {
	ole.IUnknown
}

type DropTargetPatternVtbl struct {
	ole.IUnknownVtbl
	Get_CurrentDropTargetEffect  uintptr
	Get_CachedDropTargetEffect   uintptr
	Get_CurrentDropTargetEffects uintptr
	Get_CachedDropTargetEffects  uintptr
}

func (pat *DropTargetPattern) VTable() *DropTargetPatternVtbl {
	return (*DropTargetPatternVtbl)(unsafe.Pointer(pat.RawVTable))
}

func (pat *DropTargetPattern) CurrentDropTargetEffect() (string, error) {
	var dropTargetEffect *uint16

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CurrentDropTargetEffect,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&dropTargetEffect)),
	)

	if hr != 0 {
		return "", ole.NewError(hr)
	}

	return ole.BstrToString(dropTargetEffect), nil
}

func (pat *DropTargetPattern) CachedDropTargetEffect() (string, error) {
	var dropTargetEffect *uint16

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CachedDropTargetEffect,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&dropTargetEffect)),
	)

	if hr != 0 {
		return "", ole.NewError(hr)
	}

	return ole.BstrToString(dropTargetEffect), nil
}

func (pat *DropTargetPattern) CurrentDropTargetEffects() (*ole.SafeArray, error) {
	var dropTargetEffects *ole.SafeArray

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CurrentDropTargetEffects,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&dropTargetEffects)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return dropTargetEffects, nil
}

func (pat *DropTargetPattern) CachedDropTargetEffects() (*ole.SafeArray, error) {
	var dropTargetEffects *ole.SafeArray

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Get_CachedDropTargetEffects,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&dropTargetEffects)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return dropTargetEffects, nil
}

type TextEditPattern struct {
	ole.IUnknown
}

type TextEditPatternVtbl struct {
	ole.IUnknownVtbl
	GetActiveComposition uintptr
	GetConversionTarget  uintptr
}

func (pat *TextEditPattern) VTable() *TextEditPatternVtbl {
	return (*TextEditPatternVtbl)(unsafe.Pointer(pat.RawVTable))
}

func (pat *TextEditPattern) GetActiveComposition() (*TextRange, error) {
	var composition *TextRange

	hr, _, _ := syscall.SyscallN(
		pat.VTable().GetActiveComposition,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&composition)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return composition, nil
}

func (pat *TextEditPattern) GetConversionTarget() (*TextRange, error) {
	var conversionTarget *TextRange

	hr, _, _ := syscall.SyscallN(
		pat.VTable().GetConversionTarget,
		uintptr(unsafe.Pointer(pat)),
		uintptr(unsafe.Pointer(&conversionTarget)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return conversionTarget, nil
}

type CustomNavigationPattern struct {
	ole.IUnknown
}

type CustomNavigationPatternVtbl struct {
	ole.IUnknownVtbl
	Navigate uintptr
}

func (pat *CustomNavigationPattern) VTable() *CustomNavigationPatternVtbl {
	return (*CustomNavigationPatternVtbl)(unsafe.Pointer(pat.RawVTable))
}

func (pat *CustomNavigationPattern) Navigate(direction NavigateDirection) (*Element, error) {
	var element *Element

	hr, _, _ := syscall.SyscallN(
		pat.VTable().Navigate,
		uintptr(unsafe.Pointer(pat)),
		uintptr(direction),
		uintptr(unsafe.Pointer(&element)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return element, nil
}
