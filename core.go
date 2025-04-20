package uiautomation

import (
	"syscall"
	"unsafe"

	"github.com/go-ole/go-ole"
)

type UIAutomationTreeWalker struct {
	ole.IUnknown
}

type UIAutomationTreeWalkerVtbl struct {
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

func (walker *UIAutomationTreeWalker) VTable() *UIAutomationTreeWalkerVtbl {
	return (*UIAutomationTreeWalkerVtbl)(unsafe.Pointer(walker.RawVTable))
}

type UIAutomationCacheRequest struct {
	ole.IUnknown
}

type UIAutomationCacheRequestVtbl struct {
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

func (cr *UIAutomationCacheRequest) VTable() *UIAutomationCacheRequestVtbl {
	return (*UIAutomationCacheRequestVtbl)(unsafe.Pointer(cr.RawVTable))
}

type UIAutomationProxyFactory struct {
	ole.IUnknown
}

type UIAutomationProxyFactoryVtbl struct {
	ole.IUnknownVtbl
	CreateProvider     uintptr
	Get_ProxyFactoryId uintptr
}

func (pf *UIAutomationProxyFactory) VTable() *UIAutomationProxyFactoryVtbl {
	return (*UIAutomationProxyFactoryVtbl)(unsafe.Pointer(pf.RawVTable))
}

func (pf *UIAutomationProxyFactory) CreateProvider(hwnd syscall.Handle, idObject, idChild int32) (*RawElementProviderSimple, error) {
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

func (pf *UIAutomationProxyFactory) ProxyFactoryId() (string, error) {
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

type UIAutomationProxyFactoryEntry struct {
	ole.IUnknown
}

type UIAutomationProxyFactoryEntryVtbl struct {
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

func (pfe *UIAutomationProxyFactoryEntry) VTable() *UIAutomationProxyFactoryEntryVtbl {
	return (*UIAutomationProxyFactoryEntryVtbl)(unsafe.Pointer(pfe.RawVTable))
}

type UIAutomationProxyFactoryMapping struct {
	vtbl *ole.IUnknown
}
type UIAutomationProxyFactoryMappingVtbl struct {
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

func (pfm *UIAutomationProxyFactoryMapping) VTable() *UIAutomationProxyFactoryMappingVtbl {
	return (*UIAutomationProxyFactoryMappingVtbl)(unsafe.Pointer(pfm.vtbl.RawVTable))
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
