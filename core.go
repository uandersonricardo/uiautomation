package uiautomation

import (
	"syscall"
	"unsafe"

	"github.com/go-ole/go-ole"
)

type TreeWalker struct {
	ole.IUnknown
}

type TreeWalkerVtbl struct {
	ole.IUnknownVtbl
	GetParentElement                    uintptr
	GetFirstChildElement                uintptr
	GetLastChildElement                 uintptr
	GetNextSiblingElement               uintptr
	GetPreviousSiblingElement           uintptr
	NormalizeElement                    uintptr
	GetParentElementBuildCache          uintptr
	GetFirstChildElementBuildCache      uintptr
	GetLastChildElementBuildCache       uintptr
	GetNextSiblingElementBuildCache     uintptr
	GetPreviousSiblingElementBuildCache uintptr
	NormalizeElementBuildCache          uintptr
	Get_Condition                       uintptr
}

func (walker *TreeWalker) VTable() *TreeWalkerVtbl {
	return (*TreeWalkerVtbl)(unsafe.Pointer(walker.RawVTable))
}

func (walker *TreeWalker) GetFirstChildElement(element *Element) (*Element, error) {
	var first *Element

	hr, _, _ := syscall.SyscallN(
		walker.VTable().GetFirstChildElement,
		uintptr(unsafe.Pointer(walker)),
		uintptr(unsafe.Pointer(element)),
		uintptr(unsafe.Pointer(&first)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return first, nil
}

func (walker *TreeWalker) GetNextSiblingElement(element *Element) (*Element, error) {
	var next *Element

	hr, _, _ := syscall.SyscallN(
		walker.VTable().GetNextSiblingElement,
		uintptr(unsafe.Pointer(walker)),
		uintptr(unsafe.Pointer(element)),
		uintptr(unsafe.Pointer(&next)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return next, nil
}

func (walker *TreeWalker) GetParentElement(element *Element) (*Element, error) {
	var parent *Element

	hr, _, _ := syscall.SyscallN(
		walker.VTable().GetParentElement,
		uintptr(unsafe.Pointer(walker)),
		uintptr(unsafe.Pointer(element)),
		uintptr(unsafe.Pointer(&parent)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return parent, nil
}

type CacheRequest struct {
	ole.IUnknown
}

type CacheRequestVtbl struct {
	ole.IUnknownVtbl
	AddProperty               uintptr
	AddPattern                uintptr
	Clone                     uintptr
	Get_TreeScope             uintptr
	Put_TreeScope             uintptr
	Get_TreeFilter            uintptr
	Put_TreeFilter            uintptr
	Get_AutomationElementMode uintptr
	Put_AutomationElementMode uintptr
}

func (cr *CacheRequest) VTable() *CacheRequestVtbl {
	return (*CacheRequestVtbl)(unsafe.Pointer(cr.RawVTable))
}

type ProxyFactory struct {
	ole.IUnknown
}

type ProxyFactoryVtbl struct {
	ole.IUnknownVtbl
	CreateProvider     uintptr
	Get_ProxyFactoryId uintptr
}

func (pf *ProxyFactory) VTable() *ProxyFactoryVtbl {
	return (*ProxyFactoryVtbl)(unsafe.Pointer(pf.RawVTable))
}

func (pf *ProxyFactory) CreateProvider(hwnd syscall.Handle, idObject, idChild int32) (*RawElementProviderSimple, error) {
	var provider *RawElementProviderSimple

	hr, _, _ := syscall.SyscallN(
		pf.VTable().CreateProvider,
		uintptr(unsafe.Pointer(pf)),
		uintptr(hwnd),
		uintptr(idObject),
		uintptr(idChild),
		uintptr(unsafe.Pointer(&provider)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return provider, nil
}

func (pf *ProxyFactory) ProxyFactoryId() (string, error) {
	var factoryId *uint16

	hr, _, _ := syscall.SyscallN(
		pf.VTable().Get_ProxyFactoryId,
		uintptr(unsafe.Pointer(pf)),
		uintptr(unsafe.Pointer(&factoryId)),
	)

	if hr != 0 {
		return "", ole.NewError(hr)
	}

	return ole.BstrToString(factoryId), nil
}

type ProxyFactoryEntry struct {
	ole.IUnknown
}

type ProxyFactoryEntryVtbl struct {
	ole.IUnknownVtbl
	Get_ProxyFactory               uintptr
	Get_ClassName                  uintptr
	Get_ImageName                  uintptr
	Get_AllowSubstringMatch        uintptr
	Get_CanCheckBaseClass          uintptr
	Get_NeedsAdviseEvents          uintptr
	Put_ClassName                  uintptr
	Put_ImageName                  uintptr
	Put_AllowSubstringMatch        uintptr
	Put_CanCheckBaseClass          uintptr
	Put_NeedsAdviseEvents          uintptr
	SetWinEventsForAutomationEvent uintptr
	GetWinEventsForAutomationEvent uintptr
}

func (pfe *ProxyFactoryEntry) VTable() *ProxyFactoryEntryVtbl {
	return (*ProxyFactoryEntryVtbl)(unsafe.Pointer(pfe.RawVTable))
}

type ProxyFactoryMapping struct {
	vtbl *ole.IUnknown
}
type ProxyFactoryMappingVtbl struct {
	ole.IUnknownVtbl
	Get_Count           uintptr
	GetTable            uintptr
	GetEntry            uintptr
	SetTable            uintptr
	InsertEntries       uintptr
	InsertEntry         uintptr
	RemoveEntry         uintptr
	ClearTable          uintptr
	RestoreDefaultTable uintptr
}

func (pfm *ProxyFactoryMapping) VTable() *ProxyFactoryMappingVtbl {
	return (*ProxyFactoryMappingVtbl)(unsafe.Pointer(pfm.vtbl.RawVTable))
}

type RawElementProviderSimple struct {
	ole.IUnknown
}

type RawElementProviderSimpleVtbl struct {
	ole.IUnknownVtbl
	Get_ProviderOptions        uintptr
	GetPatternProvider         uintptr
	GetPropertyValue           uintptr
	Get_HostRawElementProvider uintptr
}

func (rep *RawElementProviderSimple) VTable() *RawElementProviderSimpleVtbl {
	return (*RawElementProviderSimpleVtbl)(unsafe.Pointer(rep.RawVTable))
}

type TextRangeArray struct {
	ole.IUnknown
}

type TextRangeArrayVtbl struct {
	ole.IUnknownVtbl
	Get_Length uintptr
	GetElement uintptr
}

func (tra *TextRangeArray) VTable() *TextRangeArrayVtbl {
	return (*TextRangeArrayVtbl)(unsafe.Pointer(tra.RawVTable))
}

func (tra *TextRangeArray) Length() (int32, error) {
	var length int32

	hr, _, _ := syscall.SyscallN(
		tra.VTable().Get_Length,
		uintptr(unsafe.Pointer(tra)),
		uintptr(unsafe.Pointer(&length)),
	)

	if hr != 0 {
		return 0, ole.NewError(hr)
	}

	return length, nil
}

func (tra *TextRangeArray) GetElement(index int32) (*TextRange, error) {
	var textRange *TextRange

	hr, _, _ := syscall.SyscallN(
		tra.VTable().GetElement,
		uintptr(unsafe.Pointer(tra)),
		uintptr(index),
		uintptr(unsafe.Pointer(&textRange)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return textRange, nil
}

type TextRange struct {
	ole.IUnknown
}

type TextRangeVtbl struct {
	ole.IUnknownVtbl
	Clone                 uintptr
	Compare               uintptr
	CompareEndpoints      uintptr
	ExpandToEnclosingUnit uintptr
	FindAttribute         uintptr
	FindText              uintptr
	GetAttributeValue     uintptr
	GetBoundingRectangles uintptr
	GetEnclosingElement   uintptr
	GetText               uintptr
	Move                  uintptr
	MoveEndpointByUnit    uintptr
	MoveEndpointByRange   uintptr
	Select                uintptr
	AddToSelection        uintptr
	RemoveFromSelection   uintptr
	ScrollIntoView        uintptr
	GetChildren           uintptr
}

func (tr *TextRange) VTable() *TextRangeVtbl {
	return (*TextRangeVtbl)(unsafe.Pointer(tr.RawVTable))
}

func (tr *TextRange) GetBoundingRectangles() ([]Rect, error) {
	var arr *ole.SafeArray

	hr, _, _ := syscall.SyscallN(
		tr.VTable().GetBoundingRectangles,
		uintptr(unsafe.Pointer(tr)),
		uintptr(unsafe.Pointer(&arr)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	conversion := ole.SafeArrayConversion{Array: arr}
	defer conversion.Release()

	floats := conversion.ToValueArray()

	rects := make([]Rect, len(floats)/4)
	for i := 0; i < len(floats)/4; i++ {
		rects[i] = Rect{
			Left:   uint32(floats[i*4].(float64)),
			Top:    uint32(floats[i*4+1].(float64)),
			Right:  uint32(floats[i*4].(float64)) + uint32(floats[i*4+2].(float64)),
			Bottom: uint32(floats[i*4+1].(float64)) + uint32(floats[i*4+3].(float64)),
		}
	}

	return rects, nil
}

func (tr *TextRange) GetText(maxLength int32) (string, error) {
	var text *uint16

	hr, _, _ := syscall.SyscallN(
		tr.VTable().GetText,
		uintptr(unsafe.Pointer(tr)),
		uintptr(maxLength),
		uintptr(unsafe.Pointer(&text)),
	)

	if hr != 0 {
		return "", ole.NewError(hr)
	}

	return ole.BstrToString(text), nil
}
