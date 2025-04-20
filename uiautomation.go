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

func (auto *UIAutomation) CompareRuntimeIds(runtimeId1, runtimeId2 []int32) (bool, error) {
	panic("Not implemented")
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

func (auto *UIAutomation) ElementFromHandle(hwnd syscall.Handle) (*UIAutomationElement, error) {
	panic("Not implemented")
}

func (auto *UIAutomation) ElementFromPoint(point Point) (*UIAutomationElement, error) {
	panic("Not implemented")
}

func (auto *UIAutomation) GetFocusedElement() (*UIAutomationElement, error) {
	var element *UIAutomationElement

	hr, _, _ := syscall.SyscallN(
		auto.VTable().GetFocusedElement,
		uintptr(unsafe.Pointer(auto)),
		uintptr(unsafe.Pointer(&element)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return element, nil
}

func (auto *UIAutomation) GetRootElementBuildCache(cacheRequest *UIAutomationCacheRequest) (*UIAutomationElement, error) {
	panic("Not implemented")
}

func (auto *UIAutomation) ElementFromHandleBuildCache(hwnd syscall.Handle, cacheRequest *UIAutomationCacheRequest) (*UIAutomationElement, error) {
	panic("Not implemented")
}

func (auto *UIAutomation) ElementFromPointBuildCache(point Point, cacheRequest *UIAutomationCacheRequest) (*UIAutomationElement, error) {
	panic("Not implemented")
}

func (auto *UIAutomation) GetFocusedElementBuildCache(cacheRequest *UIAutomationCacheRequest) (*UIAutomationElement, error) {
	panic("Not implemented")
}

func (auto *UIAutomation) CreateTreeWalker(condition *UIAutomationCondition) (*UIAutomationTreeWalker, error) {
	panic("Not implemented")
}

func (auto *UIAutomation) ControlViewWalker() (*UIAutomationTreeWalker, error) {
	panic("Not implemented")
}

func (auto *UIAutomation) ContentViewWalker() (*UIAutomationTreeWalker, error) {
	panic("Not implemented")
}

func (auto *UIAutomation) RawViewWalker() (*UIAutomationTreeWalker, error) {
	panic("Not implemented")
}

func (auto *UIAutomation) RawViewCondition() (*UIAutomationCondition, error) {
	panic("Not implemented")
}

func (auto *UIAutomation) ControlViewCondition() (*UIAutomationCondition, error) {
	panic("Not implemented")
}

func (auto *UIAutomation) ContentViewCondition() (*UIAutomationCondition, error) {
	panic("Not implemented")
}

func (auto *UIAutomation) CreateCacheRequest() (*UIAutomationCacheRequest, error) {
	panic("Not implemented")
}

func (auto *UIAutomation) CreateTrueCondition() (*UIAutomationCondition, error) {
	panic("Not implemented")
}

func (auto *UIAutomation) CreateFalseCondition() (*UIAutomationCondition, error) {
	panic("Not implemented")
}

func (auto *UIAutomation) CreatePropertyCondition(propertyId int32, value interface{}) (*UIAutomationCondition, error) {
	panic("Not implemented")
}

func (auto *UIAutomation) CreatePropertyConditionEx(propertyId int32, value interface{}, flags PropertyConditionFlags) (*UIAutomationCondition, error) {
	panic("Not implemented")
}

func (auto *UIAutomation) CreateAndCondition(condition1, condition2 *UIAutomationCondition) (*UIAutomationCondition, error) {
	panic("Not implemented")
}

func (auto *UIAutomation) CreateAndConditionFromArray(conditions []*UIAutomationCondition) (*UIAutomationCondition, error) {
	panic("Not implemented")
}

func (auto *UIAutomation) CreateAndConditionFromNativeArray(conditions []int32) (*UIAutomationCondition, error) {
	panic("Not implemented")
}

func (auto *UIAutomation) CreateOrCondition(condition1, condition2 *UIAutomationCondition) (*UIAutomationCondition, error) {
	panic("Not implemented")
}

func (auto *UIAutomation) CreateOrConditionFromArray(conditions []*UIAutomationCondition) (*UIAutomationCondition, error) {
	panic("Not implemented")
}

func (auto *UIAutomation) CreateOrConditionFromNativeArray(conditions []int32) (*UIAutomationCondition, error) {
	panic("Not implemented")
}

func (auto *UIAutomation) CreateNotCondition(condition *UIAutomationCondition) (*UIAutomationCondition, error) {
	panic("Not implemented")
}

func (auto *UIAutomation) AddAutomationEventHandler(eventId EventId, element *UIAutomationElement, scope TreeScope, cacheRequest *UIAutomationCacheRequest, handler *UIAutomationEventHandler) error {
	panic("Not implemented")
}

func (auto *UIAutomation) RemoveAutomationEventHandler(eventId EventId, element *UIAutomationElement, handler *UIAutomationEventHandler) error {
	panic("Not implemented")
}

func (auto *UIAutomation) AddPropertyChangedEventHandlerNativeArray(element *UIAutomationElement, scope TreeScope, cacheRequest *UIAutomationCacheRequest, handler *UIAutomationPropertyChangedEventHandler, propertyArray []int32) error {
	panic("Not implemented")
}

func (auto *UIAutomation) AddPropertyChangedEventHandler(element *UIAutomationElement, scope TreeScope, cacheRequest *UIAutomationCacheRequest, handler *UIAutomationPropertyChangedEventHandler, propertyArray []PropertyId) error {
	panic("Not implemented")
}

func (auto *UIAutomation) RemovePropertyChangedEventHandler(element *UIAutomationElement, handler *UIAutomationPropertyChangedEventHandler) error {
	panic("Not implemented")
}

func (auto *UIAutomation) AddStructureChangedEventHandler(element *UIAutomationElement, scope TreeScope, cacheRequest *UIAutomationCacheRequest, handler *UIAutomationStructureChangedEventHandler) error {
	panic("Not implemented")
}

func (auto *UIAutomation) RemoveStructureChangedEventHandler(element *UIAutomationElement, handler *UIAutomationStructureChangedEventHandler) error {
	panic("Not implemented")
}

func (auto *UIAutomation) AddFocusChangedEventHandler(handler *UIAutomationFocusChangedEventHandler) error {
	panic("Not implemented")
}

func (auto *UIAutomation) RemoveFocusChangedEventHandler(handler *UIAutomationFocusChangedEventHandler) error {
	panic("Not implemented")
}

func (auto *UIAutomation) RemoveAllEventHandlers() error {
	panic("Not implemented")
}

func (auto *UIAutomation) IntNativeArrayToSafeArray(array []int32) (*ole.SafeArray, error) {
	panic("Not implemented")
}

func (auto *UIAutomation) IntSafeArrayToNativeArray(safeArray *ole.SafeArray) ([]int32, error) {
	panic("Not implemented")
}

func (auto *UIAutomation) RectToVariant(rect Rect) (ole.VARIANT, error) {
	panic("Not implemented")
}

func (auto *UIAutomation) VariantToRect(variant ole.VARIANT) (Rect, error) {
	panic("Not implemented")
}

func (auto *UIAutomation) SafeArrayToRectNativeArray(safeArray *ole.SafeArray) ([]Rect, error) {
	panic("Not implemented")
}

func (auto *UIAutomation) CreateProxyFactoryEntry(factory *UIAutomationProxyFactory) (*UIAutomationProxyFactoryEntry, error) {
	panic("Not implemented")
}

func (auto *UIAutomation) ProxyFactoryMapping() (*UIAutomationProxyFactoryMapping, error) {
	panic("Not implemented")
}

func (auto *UIAutomation) GetPropertyProgrammaticName(propertyId PropertyId) (string, error) {
	panic("Not implemented")
}

func (auto *UIAutomation) GetPatternProgrammaticName(patternId PatternId) (string, error) {
	panic("Not implemented")
}

func (auto *UIAutomation) PollForPotentialSupportedPatterns(element *UIAutomationElement) ([]PatternId, error) {
	panic("Not implemented")
}

func (auto *UIAutomation) PollForPotentialSupportedProperties(element *UIAutomationElement) ([]PropertyId, error) {
	panic("Not implemented")
}

func (auto *UIAutomation) CheckNotSupported(value interface{}) (bool, error) {
	panic("Not implemented")
}

func (auto *UIAutomation) ReservedNotSupportedValue() (interface{}, error) {
	panic("Not implemented")
}

func (auto *UIAutomation) ReservedMixedAttributeValue() (interface{}, error) {
	panic("Not implemented")
}

func (auto *UIAutomation) ElementFromIAccessible(accessible *Accessible, childId int32) (*UIAutomationElement, error) {
	panic("Not implemented")
}

func (auto *UIAutomation) ElementFromIAccessibleBuildCache(accessible *Accessible, childId int32, cacheRequest *UIAutomationCacheRequest) (*UIAutomationElement, error) {
	panic("Not implemented")
}
