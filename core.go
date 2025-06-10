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

	if parent == nil {
		return nil, ole.NewError(ole.E_POINTER)
	}

	return parent, nil
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

	if first == nil {
		return nil, ole.NewError(ole.E_POINTER)
	}

	return first, nil
}

func (walker *TreeWalker) GetLastChildElement(element *Element) (*Element, error) {
	var last *Element

	hr, _, _ := syscall.SyscallN(
		walker.VTable().GetLastChildElement,
		uintptr(unsafe.Pointer(walker)),
		uintptr(unsafe.Pointer(element)),
		uintptr(unsafe.Pointer(&last)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	if last == nil {
		return nil, ole.NewError(ole.E_POINTER)
	}

	return last, nil
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

	if next == nil {
		return nil, ole.NewError(ole.E_POINTER)
	}

	return next, nil
}

func (walker *TreeWalker) GetPreviousSiblingElement(element *Element) (*Element, error) {
	var previous *Element

	hr, _, _ := syscall.SyscallN(
		walker.VTable().GetPreviousSiblingElement,
		uintptr(unsafe.Pointer(walker)),
		uintptr(unsafe.Pointer(element)),
		uintptr(unsafe.Pointer(&previous)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	if previous == nil {
		return nil, ole.NewError(ole.E_POINTER)
	}

	return previous, nil
}

func (walker *TreeWalker) NormalizeElement(element *Element) (*Element, error) {
	var normalized *Element

	hr, _, _ := syscall.SyscallN(
		walker.VTable().NormalizeElement,
		uintptr(unsafe.Pointer(walker)),
		uintptr(unsafe.Pointer(element)),
		uintptr(unsafe.Pointer(&normalized)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	if normalized == nil {
		return nil, ole.NewError(ole.E_POINTER)
	}

	return normalized, nil
}

func (walker *TreeWalker) GetParentElementBuildCache(element *Element, cacheRequest *CacheRequest) (*Element, error) {
	var parent *Element

	hr, _, _ := syscall.SyscallN(
		walker.VTable().GetParentElementBuildCache,
		uintptr(unsafe.Pointer(walker)),
		uintptr(unsafe.Pointer(element)),
		uintptr(unsafe.Pointer(cacheRequest)),
		uintptr(unsafe.Pointer(&parent)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	if parent == nil {
		return nil, ole.NewError(ole.E_POINTER)
	}

	return parent, nil
}

func (walker *TreeWalker) GetFirstChildElementBuildCache(element *Element, cacheRequest *CacheRequest) (*Element, error) {
	var first *Element

	hr, _, _ := syscall.SyscallN(
		walker.VTable().GetFirstChildElementBuildCache,
		uintptr(unsafe.Pointer(walker)),
		uintptr(unsafe.Pointer(element)),
		uintptr(unsafe.Pointer(cacheRequest)),
		uintptr(unsafe.Pointer(&first)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	if first == nil {
		return nil, ole.NewError(ole.E_POINTER)
	}

	return first, nil
}

func (walker *TreeWalker) GetLastChildElementBuildCache(element *Element, cacheRequest *CacheRequest) (*Element, error) {
	var last *Element

	hr, _, _ := syscall.SyscallN(
		walker.VTable().GetLastChildElementBuildCache,
		uintptr(unsafe.Pointer(walker)),
		uintptr(unsafe.Pointer(element)),
		uintptr(unsafe.Pointer(cacheRequest)),
		uintptr(unsafe.Pointer(&last)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	if last == nil {
		return nil, ole.NewError(ole.E_POINTER)
	}

	return last, nil
}

func (walker *TreeWalker) GetNextSiblingElementBuildCache(element *Element, cacheRequest *CacheRequest) (*Element, error) {
	var next *Element

	hr, _, _ := syscall.SyscallN(
		walker.VTable().GetNextSiblingElementBuildCache,
		uintptr(unsafe.Pointer(walker)),
		uintptr(unsafe.Pointer(element)),
		uintptr(unsafe.Pointer(cacheRequest)),
		uintptr(unsafe.Pointer(&next)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	if next == nil {
		return nil, ole.NewError(ole.E_POINTER)
	}

	return next, nil
}

func (walker *TreeWalker) GetPreviousSiblingElementBuildCache(element *Element, cacheRequest *CacheRequest) (*Element, error) {
	var previous *Element

	hr, _, _ := syscall.SyscallN(
		walker.VTable().GetPreviousSiblingElementBuildCache,
		uintptr(unsafe.Pointer(walker)),
		uintptr(unsafe.Pointer(element)),
		uintptr(unsafe.Pointer(cacheRequest)),
		uintptr(unsafe.Pointer(&previous)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	if previous == nil {
		return nil, ole.NewError(ole.E_POINTER)
	}

	return previous, nil
}

func (walker *TreeWalker) NormalizeElementBuildCache(element *Element, cacheRequest *CacheRequest) (*Element, error) {
	var normalized *Element

	hr, _, _ := syscall.SyscallN(
		walker.VTable().NormalizeElementBuildCache,
		uintptr(unsafe.Pointer(walker)),
		uintptr(unsafe.Pointer(element)),
		uintptr(unsafe.Pointer(cacheRequest)),
		uintptr(unsafe.Pointer(&normalized)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	if normalized == nil {
		return nil, ole.NewError(ole.E_POINTER)
	}

	return normalized, nil
}

func (walker *TreeWalker) Condition() (*Condition, error) {
	var condition *Condition

	hr, _, _ := syscall.SyscallN(
		walker.VTable().Get_Condition,
		uintptr(unsafe.Pointer(walker)),
		uintptr(unsafe.Pointer(&condition)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return condition, nil
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

func (cr *CacheRequest) AddProperty(propertyId PropertyId) error {
	hr, _, _ := syscall.SyscallN(
		cr.VTable().AddProperty,
		uintptr(unsafe.Pointer(cr)),
		uintptr(propertyId),
	)

	if hr != 0 {
		return ole.NewError(hr)
	}

	return nil
}

func (cr *CacheRequest) AddPattern(patternId PatternId) error {
	hr, _, _ := syscall.SyscallN(
		cr.VTable().AddPattern,
		uintptr(unsafe.Pointer(cr)),
		uintptr(patternId),
	)

	if hr != 0 {
		return ole.NewError(hr)
	}

	return nil
}

func (cr *CacheRequest) Clone() (*CacheRequest, error) {
	var clone *CacheRequest

	hr, _, _ := syscall.SyscallN(
		cr.VTable().Clone,
		uintptr(unsafe.Pointer(cr)),
		uintptr(unsafe.Pointer(&clone)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return clone, nil
}

func (cr *CacheRequest) TreeScope() (TreeScope, error) {
	var scope TreeScope

	hr, _, _ := syscall.SyscallN(
		cr.VTable().Get_TreeScope,
		uintptr(unsafe.Pointer(cr)),
		uintptr(unsafe.Pointer(&scope)),
	)

	if hr != 0 {
		return 0, ole.NewError(hr)
	}

	return scope, nil
}

func (cr *CacheRequest) SetTreeScope(scope TreeScope) error {
	hr, _, _ := syscall.SyscallN(
		cr.VTable().Put_TreeScope,
		uintptr(unsafe.Pointer(cr)),
		uintptr(scope),
	)

	if hr != 0 {
		return ole.NewError(hr)
	}

	return nil
}

func (cr *CacheRequest) TreeFilter() (*Condition, error) {
	var filter *Condition

	hr, _, _ := syscall.SyscallN(
		cr.VTable().Get_TreeFilter,
		uintptr(unsafe.Pointer(cr)),
		uintptr(unsafe.Pointer(&filter)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return filter, nil
}

func (cr *CacheRequest) SetTreeFilter(filter *Condition) error {
	hr, _, _ := syscall.SyscallN(
		cr.VTable().Put_TreeFilter,
		uintptr(unsafe.Pointer(cr)),
		uintptr(unsafe.Pointer(filter)),
	)

	if hr != 0 {
		return ole.NewError(hr)
	}

	return nil
}

func (cr *CacheRequest) AutomationElementMode() (AutomationElementMode, error) {
	var mode AutomationElementMode

	hr, _, _ := syscall.SyscallN(
		cr.VTable().Get_AutomationElementMode,
		uintptr(unsafe.Pointer(cr)),
		uintptr(unsafe.Pointer(&mode)),
	)

	if hr != 0 {
		return 0, ole.NewError(hr)
	}

	return mode, nil
}

func (cr *CacheRequest) SetAutomationElementMode(mode AutomationElementMode) error {
	hr, _, _ := syscall.SyscallN(
		cr.VTable().Put_AutomationElementMode,
		uintptr(unsafe.Pointer(cr)),
		uintptr(mode),
	)

	if hr != 0 {
		return ole.NewError(hr)
	}

	return nil
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

func (pfe *ProxyFactoryEntry) ProxyFactory() (*ProxyFactory, error) {
	var factory *ProxyFactory

	hr, _, _ := syscall.SyscallN(
		pfe.VTable().Get_ProxyFactory,
		uintptr(unsafe.Pointer(pfe)),
		uintptr(unsafe.Pointer(&factory)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return factory, nil
}

func (pfe *ProxyFactoryEntry) ClassName() (string, error) {
	var className *uint16

	hr, _, _ := syscall.SyscallN(
		pfe.VTable().Get_ClassName,
		uintptr(unsafe.Pointer(pfe)),
		uintptr(unsafe.Pointer(&className)),
	)

	if hr != 0 {
		return "", ole.NewError(hr)
	}

	return ole.BstrToString(className), nil
}

func (pfe *ProxyFactoryEntry) ImageName() (string, error) {
	var imageName *uint16

	hr, _, _ := syscall.SyscallN(
		pfe.VTable().Get_ImageName,
		uintptr(unsafe.Pointer(pfe)),
		uintptr(unsafe.Pointer(&imageName)),
	)

	if hr != 0 {
		return "", ole.NewError(hr)
	}

	return ole.BstrToString(imageName), nil
}

func (pfe *ProxyFactoryEntry) AllowSubstringMatch() (bool, error) {
	var allow bool

	hr, _, _ := syscall.SyscallN(
		pfe.VTable().Get_AllowSubstringMatch,
		uintptr(unsafe.Pointer(pfe)),
		uintptr(unsafe.Pointer(&allow)),
	)

	if hr != 0 {
		return false, ole.NewError(hr)
	}

	return allow, nil
}

func (pfe *ProxyFactoryEntry) CanCheckBaseClass() (bool, error) {
	var canCheck bool

	hr, _, _ := syscall.SyscallN(
		pfe.VTable().Get_CanCheckBaseClass,
		uintptr(unsafe.Pointer(pfe)),
		uintptr(unsafe.Pointer(&canCheck)),
	)

	if hr != 0 {
		return false, ole.NewError(hr)
	}

	return canCheck, nil
}

func (pfe *ProxyFactoryEntry) NeedsAdviseEvents() (bool, error) {
	var needsAdvise bool

	hr, _, _ := syscall.SyscallN(
		pfe.VTable().Get_NeedsAdviseEvents,
		uintptr(unsafe.Pointer(pfe)),
		uintptr(unsafe.Pointer(&needsAdvise)),
	)

	if hr != 0 {
		return false, ole.NewError(hr)
	}

	return needsAdvise, nil
}

func (pfe *ProxyFactoryEntry) SetClassName(className string) error {
	bstr := ole.SysAllocString(className)
	defer ole.SysFreeString(bstr)

	hr, _, _ := syscall.SyscallN(
		pfe.VTable().Put_ClassName,
		uintptr(unsafe.Pointer(pfe)),
		uintptr(unsafe.Pointer(bstr)),
	)

	if hr != 0 {
		return ole.NewError(hr)
	}

	return nil
}

func (pfe *ProxyFactoryEntry) SetImageName(imageName string) error {
	bstr := ole.SysAllocString(imageName)
	defer ole.SysFreeString(bstr)

	hr, _, _ := syscall.SyscallN(
		pfe.VTable().Put_ImageName,
		uintptr(unsafe.Pointer(pfe)),
		uintptr(unsafe.Pointer(bstr)),
	)

	if hr != 0 {
		return ole.NewError(hr)
	}

	return nil
}

func (pfe *ProxyFactoryEntry) SetAllowSubstringMatch(allow bool) error {
	var value int32 = 0

	if allow {
		value = 1
	}

	hr, _, _ := syscall.SyscallN(
		pfe.VTable().Put_AllowSubstringMatch,
		uintptr(unsafe.Pointer(pfe)),
		uintptr(value),
	)

	if hr != 0 {
		return ole.NewError(hr)
	}

	return nil
}

func (pfe *ProxyFactoryEntry) SetCanCheckBaseClass(canCheck bool) error {
	var value int32 = 0

	if canCheck {
		value = 1
	}

	hr, _, _ := syscall.SyscallN(
		pfe.VTable().Put_CanCheckBaseClass,
		uintptr(unsafe.Pointer(pfe)),
		uintptr(value),
	)

	if hr != 0 {
		return ole.NewError(hr)
	}

	return nil
}

func (pfe *ProxyFactoryEntry) SetNeedsAdviseEvents(needsAdvise bool) error {
	var value int32 = 0

	if needsAdvise {
		value = 1
	}

	hr, _, _ := syscall.SyscallN(
		pfe.VTable().Put_NeedsAdviseEvents,
		uintptr(unsafe.Pointer(pfe)),
		uintptr(value),
	)

	if hr != 0 {
		return ole.NewError(hr)
	}

	return nil
}

type ProxyFactoryMapping struct {
	ole.IUnknown
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
	return (*ProxyFactoryMappingVtbl)(unsafe.Pointer(pfm.RawVTable))
}

func (pfm *ProxyFactoryMapping) Count() (int32, error) {
	var count int32

	hr, _, _ := syscall.SyscallN(
		pfm.VTable().Get_Count,
		uintptr(unsafe.Pointer(pfm)),
		uintptr(unsafe.Pointer(&count)),
	)

	if hr != 0 {
		return 0, ole.NewError(hr)
	}

	return count, nil
}

func (pfm *ProxyFactoryMapping) GetTable() (*ole.SafeArray, error) {
	var table *ole.SafeArray

	hr, _, _ := syscall.SyscallN(
		pfm.VTable().GetTable,
		uintptr(unsafe.Pointer(pfm)),
		uintptr(unsafe.Pointer(&table)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return table, nil
}

func (pfm *ProxyFactoryMapping) GetEntry(index int32) (*ProxyFactoryEntry, error) {
	var entry *ProxyFactoryEntry

	hr, _, _ := syscall.SyscallN(
		pfm.VTable().GetEntry,
		uintptr(unsafe.Pointer(pfm)),
		uintptr(index),
		uintptr(unsafe.Pointer(&entry)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return entry, nil
}

func (pfm *ProxyFactoryMapping) SetTable(entries *ole.SafeArray) error {
	hr, _, _ := syscall.SyscallN(
		pfm.VTable().SetTable,
		uintptr(unsafe.Pointer(pfm)),
		uintptr(unsafe.Pointer(entries)),
	)

	if hr != 0 {
		return ole.NewError(hr)
	}

	return nil
}

func (pfm *ProxyFactoryMapping) InsertEntries(entries *ole.SafeArray) error {
	hr, _, _ := syscall.SyscallN(
		pfm.VTable().InsertEntries,
		uintptr(unsafe.Pointer(pfm)),
		uintptr(unsafe.Pointer(entries)),
	)

	if hr != 0 {
		return ole.NewError(hr)
	}

	return nil
}

func (pfm *ProxyFactoryMapping) InsertEntry(entry *ProxyFactoryEntry) error {
	hr, _, _ := syscall.SyscallN(
		pfm.VTable().InsertEntry,
		uintptr(unsafe.Pointer(pfm)),
		uintptr(unsafe.Pointer(entry)),
	)

	if hr != 0 {
		return ole.NewError(hr)
	}

	return nil
}

func (pfm *ProxyFactoryMapping) RemoveEntry(entry *ProxyFactoryEntry) error {
	hr, _, _ := syscall.SyscallN(
		pfm.VTable().RemoveEntry,
		uintptr(unsafe.Pointer(pfm)),
		uintptr(unsafe.Pointer(entry)),
	)

	if hr != 0 {
		return ole.NewError(hr)
	}

	return nil
}

func (pfm *ProxyFactoryMapping) ClearTable() error {
	hr, _, _ := syscall.SyscallN(
		pfm.VTable().ClearTable,
		uintptr(unsafe.Pointer(pfm)),
	)

	if hr != 0 {
		return ole.NewError(hr)
	}

	return nil
}

func (pfm *ProxyFactoryMapping) RestoreDefaultTable() error {
	hr, _, _ := syscall.SyscallN(
		pfm.VTable().RestoreDefaultTable,
		uintptr(unsafe.Pointer(pfm)),
	)

	if hr != 0 {
		return ole.NewError(hr)
	}

	return nil
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

func (rep *RawElementProviderSimple) ProviderOptions() (ProviderOptions, error) {
	var options ProviderOptions

	hr, _, _ := syscall.SyscallN(
		rep.VTable().Get_ProviderOptions,
		uintptr(unsafe.Pointer(rep)),
		uintptr(unsafe.Pointer(&options)),
	)

	if hr != 0 {
		return 0, ole.NewError(hr)
	}

	return options, nil
}

func (rep *RawElementProviderSimple) GetPatternProvider(patternId PatternId) (*ole.IUnknown, error) {
	var pattern *ole.IUnknown

	hr, _, _ := syscall.SyscallN(
		rep.VTable().GetPatternProvider,
		uintptr(unsafe.Pointer(rep)),
		uintptr(patternId),
		uintptr(unsafe.Pointer(&pattern)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return pattern, nil
}

func (rep *RawElementProviderSimple) GetPropertyValue(propertyId PropertyId) (*ole.VARIANT, error) {
	value := &ole.VARIANT{}

	ole.VariantInit(value)

	hr, _, _ := syscall.SyscallN(
		rep.VTable().GetPropertyValue,
		uintptr(unsafe.Pointer(rep)),
		uintptr(propertyId),
		uintptr(unsafe.Pointer(value)),
	)

	if hr != 0 {
		ole.VariantClear(value)
		return nil, ole.NewError(hr)
	}

	return value, nil
}

func (rep *RawElementProviderSimple) GetHostRawElementProvider() (*RawElementProviderSimple, error) {
	var host *RawElementProviderSimple

	hr, _, _ := syscall.SyscallN(
		rep.VTable().Get_HostRawElementProvider,
		uintptr(unsafe.Pointer(rep)),
		uintptr(unsafe.Pointer(&host)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return host, nil
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

func (tr *TextRange) Clone() (*TextRange, error) {
	var clone *TextRange

	hr, _, _ := syscall.SyscallN(
		tr.VTable().Clone,
		uintptr(unsafe.Pointer(tr)),
		uintptr(unsafe.Pointer(&clone)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return clone, nil
}

func (tr *TextRange) Compare(rangeToCompare *TextRange) (bool, error) {
	var result bool

	hr, _, _ := syscall.SyscallN(
		tr.VTable().Compare,
		uintptr(unsafe.Pointer(tr)),
		uintptr(unsafe.Pointer(rangeToCompare)),
		uintptr(unsafe.Pointer(&result)),
	)

	if hr != 0 {
		return false, ole.NewError(hr)
	}

	return result, nil
}

func (tr *TextRange) CompareEndpoints(endpoint TextPatternRangeEndpoint, rangeToCompare *TextRange, targetEndpoint TextPatternRangeEndpoint) (int32, error) {
	var result int32

	hr, _, _ := syscall.SyscallN(
		tr.VTable().CompareEndpoints,
		uintptr(unsafe.Pointer(tr)),
		uintptr(endpoint),
		uintptr(unsafe.Pointer(rangeToCompare)),
		uintptr(targetEndpoint),
		uintptr(unsafe.Pointer(&result)),
	)

	if hr != 0 {
		return 0, ole.NewError(hr)
	}

	return result, nil
}

func (tr *TextRange) ExpandToEnclosingUnit(unit TextUnit) error {
	hr, _, _ := syscall.SyscallN(
		tr.VTable().ExpandToEnclosingUnit,
		uintptr(unsafe.Pointer(tr)),
		uintptr(unit),
	)

	if hr != 0 {
		return ole.NewError(hr)
	}

	return nil
}

func (tr *TextRange) FindAttribute(attributeId AttributeId, value *ole.VARIANT, backward bool) (*TextRange, error) {
	var foundRange *TextRange

	var backwardValue int32 = 0

	if backward {
		backwardValue = 1
	}

	hr, _, _ := syscall.SyscallN(
		tr.VTable().FindAttribute,
		uintptr(unsafe.Pointer(tr)),
		uintptr(attributeId),
		uintptr(unsafe.Pointer(value)),
		uintptr(backwardValue),
		uintptr(unsafe.Pointer(&foundRange)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return foundRange, nil
}

func (tr *TextRange) FindText(text string, backward, ignoreCase bool) (*TextRange, error) {
	var foundRange *TextRange
	bstr := ole.SysAllocString(text)
	defer ole.SysFreeString(bstr)

	var backwardValue int32 = 0

	if backward {
		backwardValue = 1
	}

	var ignoreCaseValue int32 = 0

	if ignoreCase {
		ignoreCaseValue = 1
	}

	hr, _, _ := syscall.SyscallN(
		tr.VTable().FindText,
		uintptr(unsafe.Pointer(tr)),
		uintptr(unsafe.Pointer(bstr)),
		uintptr(backwardValue),
		uintptr(ignoreCaseValue),
		uintptr(unsafe.Pointer(&foundRange)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return foundRange, nil
}

func (tr *TextRange) GetAttributeValue(attributeId AttributeId) (*ole.VARIANT, error) {
	value := &ole.VARIANT{}

	ole.VariantInit(value)

	hr, _, _ := syscall.SyscallN(
		tr.VTable().GetAttributeValue,
		uintptr(unsafe.Pointer(tr)),
		uintptr(attributeId),
		uintptr(unsafe.Pointer(value)),
	)

	if hr != 0 {
		ole.VariantClear(value)
		return nil, ole.NewError(hr)
	}

	return value, nil
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

func (tr *TextRange) GetEnclosingElement() (*Element, error) {
	var element *Element

	hr, _, _ := syscall.SyscallN(
		tr.VTable().GetEnclosingElement,
		uintptr(unsafe.Pointer(tr)),
		uintptr(unsafe.Pointer(&element)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return element, nil
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

func (tr *TextRange) Move(unit TextUnit, count int32) (int32, error) {
	var moved int32

	hr, _, _ := syscall.SyscallN(
		tr.VTable().Move,
		uintptr(unsafe.Pointer(tr)),
		uintptr(unit),
		uintptr(count),
		uintptr(unsafe.Pointer(&moved)),
	)

	if hr != 0 {
		return 0, ole.NewError(hr)
	}

	return moved, nil
}

func (tr *TextRange) MoveEndpointByUnit(endpoint TextPatternRangeEndpoint, unit TextUnit, count int32) (int32, error) {
	var moved int32

	hr, _, _ := syscall.SyscallN(
		tr.VTable().MoveEndpointByUnit,
		uintptr(unsafe.Pointer(tr)),
		uintptr(endpoint),
		uintptr(unit),
		uintptr(count),
		uintptr(unsafe.Pointer(&moved)),
	)

	if hr != 0 {
		return 0, ole.NewError(hr)
	}

	return moved, nil
}

func (tr *TextRange) MoveEndpointByRange(endpoint TextPatternRangeEndpoint, targetRange *TextRange, targetEndpoint TextPatternRangeEndpoint) error {
	hr, _, _ := syscall.SyscallN(
		tr.VTable().MoveEndpointByRange,
		uintptr(unsafe.Pointer(tr)),
		uintptr(endpoint),
		uintptr(unsafe.Pointer(targetRange)),
		uintptr(targetEndpoint),
	)

	if hr != 0 {
		return ole.NewError(hr)
	}

	return nil
}

func (tr *TextRange) Select() error {
	hr, _, _ := syscall.SyscallN(
		tr.VTable().Select,
		uintptr(unsafe.Pointer(tr)),
	)

	if hr != 0 {
		return ole.NewError(hr)
	}

	return nil
}

func (tr *TextRange) AddToSelection() error {
	hr, _, _ := syscall.SyscallN(
		tr.VTable().AddToSelection,
		uintptr(unsafe.Pointer(tr)),
	)

	if hr != 0 {
		return ole.NewError(hr)
	}

	return nil
}

func (tr *TextRange) RemoveFromSelection() error {
	hr, _, _ := syscall.SyscallN(
		tr.VTable().RemoveFromSelection,
		uintptr(unsafe.Pointer(tr)),
	)

	if hr != 0 {
		return ole.NewError(hr)
	}

	return nil
}

func (tr *TextRange) ScrollIntoView(alignToTop bool) error {
	var value int32 = 0

	if alignToTop {
		value = 1
	}

	hr, _, _ := syscall.SyscallN(
		tr.VTable().ScrollIntoView,
		uintptr(unsafe.Pointer(tr)),
		uintptr(value),
	)

	if hr != 0 {
		return ole.NewError(hr)
	}

	return nil
}

func (tr *TextRange) GetChildren() (*TextRangeArray, error) {
	var children *TextRangeArray

	hr, _, _ := syscall.SyscallN(
		tr.VTable().GetChildren,
		uintptr(unsafe.Pointer(tr)),
		uintptr(unsafe.Pointer(&children)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return children, nil
}
