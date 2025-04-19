package uiautomation

import (
	"syscall"
	"unsafe"

	"github.com/go-ole/go-ole"
)

type UIAutomation struct {
	ole.IUnknown
}

type UIAutomationVtbl struct {
	ole.IUnknownVtbl
	CompareElements                           uintptr
	CompareRuntimeIds                         uintptr
	GetRootElement                            uintptr
	ElementFromHandle                         uintptr
	ElementFromPoint                          uintptr
	GetFocusedElement                         uintptr
	GetRootElementBuildCache                  uintptr
	ElementFromHandleBuildCache               uintptr
	ElementFromPointBuildCache                uintptr
	GetFocusedElementBuildCache               uintptr
	CreateTreeWalker                          uintptr
	Get_ControlViewWalker                     uintptr
	Get_ContentViewWalker                     uintptr
	Get_RawViewWalker                         uintptr
	Get_RawViewCondition                      uintptr
	Get_ControlViewCondition                  uintptr
	Get_ContentViewCondition                  uintptr
	CreateCacheRequest                        uintptr
	CreateTrueCondition                       uintptr
	CreateFalseCondition                      uintptr
	CreatePropertyCondition                   uintptr
	CreatePropertyConditionEx                 uintptr
	CreateAndCondition                        uintptr
	CreateAndConditionFromArray               uintptr
	CreateAndConditionFromNativeArray         uintptr
	CreateOrCondition                         uintptr
	CreateOrConditionFromArray                uintptr
	CreateOrConditionFromNativeArray          uintptr
	CreateNotCondition                        uintptr
	AddAutomationEventHandler                 uintptr
	RemoveAutomationEventHandler              uintptr
	AddPropertyChangedEventHandlerNativeArray uintptr
	AddPropertyChangedEventHandler            uintptr
	RemovePropertyChangedEventHandler         uintptr
	AddStructureChangedEventHandler           uintptr
	RemoveStructureChangedEventHandler        uintptr
	AddFocusChangedEventHandler               uintptr
	RemoveFocusChangedEventHandler            uintptr
	RemoveAllEventHandlers                    uintptr
	IntNativeArrayToSafeArray                 uintptr
	IntSafeArrayToNativeArray                 uintptr
	RectToVariant                             uintptr
	VariantToRect                             uintptr
	SafeArrayToRectNativeArray                uintptr
	CreateProxyFactoryEntry                   uintptr
	Get_ProxyFactoryMapping                   uintptr
	GetPropertyProgrammaticName               uintptr
	GetPatternProgrammaticName                uintptr
	PollForPotentialSupportedPatterns         uintptr
	PollForPotentialSupportedProperties       uintptr
	CheckNotSupported                         uintptr
	Get_ReservedNotSupportedValue             uintptr
	Get_ReservedMixedAttributeValue           uintptr
	ElementFromIAccessible                    uintptr
	ElementFromIAccessibleBuildCache          uintptr
}

func (auto *UIAutomation) VTable() *UIAutomationVtbl {
	return (*UIAutomationVtbl)(unsafe.Pointer(auto.RawVTable))
}

func NewUIAutomation() (*UIAutomation, error) {
	auto, err := ole.CreateInstance(CLSID_CUIAutomation, IID_IUIAutomation)

	if err != nil {
		return nil, err
	}

	return (*UIAutomation)(unsafe.Pointer(auto)), nil
}

func (auto *UIAutomation) CompareElements(el1, el2 *UIAutomation) (bool, error) {
	var areSame bool

	hr, _, _ := syscall.SyscallN(
		auto.VTable().CompareElements,
		uintptr(unsafe.Pointer(auto)),
		uintptr(unsafe.Pointer(el1)),
		uintptr(unsafe.Pointer(el2)),
		uintptr(unsafe.Pointer(&areSame)),
	)

	if hr != 0 {
		return false, ole.NewError(hr)
	}

	return areSame, nil
}

func (auto *UIAutomation) GetRootElement() (*UIAutomationElement, error) {
	var root *UIAutomationElement

	hr, _, _ := syscall.SyscallN(
		auto.VTable().GetRootElement,
		uintptr(unsafe.Pointer(auto)),
		uintptr(unsafe.Pointer(&root)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return root, nil
}

func (auto *UIAutomation) GetFocusedElement() (*UIAutomationElement, error) {
	var element *UIAutomationElement

	hr, _, _ := syscall.SyscallN(
		auto.VTable().GetRootElement,
		uintptr(unsafe.Pointer(auto)),
		uintptr(unsafe.Pointer(&element)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return element, nil
}
