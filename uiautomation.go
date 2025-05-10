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
	var areSame bool

	sa1, err := auto.IntNativeArrayToSafeArray(runtimeId1)

	if err != nil {
		return false, err
	}

	sa2, err := auto.IntNativeArrayToSafeArray(runtimeId2)

	if err != nil {
		return false, err
	}

	hr, _, _ := syscall.SyscallN(
		auto.VTable().CompareRuntimeIds,
		uintptr(unsafe.Pointer(auto)),
		uintptr(unsafe.Pointer(sa1)),
		uintptr(unsafe.Pointer(sa2)),
		uintptr(unsafe.Pointer(&areSame)),
	)

	if hr != 0 {
		return false, ole.NewError(hr)
	}

	return areSame, nil
}

func (auto *UIAutomation) GetRootElement() (*Element, error) {
	var root *Element

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

func (auto *UIAutomation) ElementFromHandle(hwnd syscall.Handle) (*Element, error) {
	var element *Element

	hr, _, _ := syscall.SyscallN(
		auto.VTable().ElementFromHandle,
		uintptr(unsafe.Pointer(auto)),
		uintptr(hwnd),
		uintptr(unsafe.Pointer(&element)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return element, nil
}

func (auto *UIAutomation) ElementFromPoint(point *Point) (*Element, error) {
	var element *Element

	hr, _, _ := syscall.SyscallN(
		auto.VTable().ElementFromPoint,
		uintptr(unsafe.Pointer(auto)),
		uintptr(unsafe.Pointer(point)),
		uintptr(unsafe.Pointer(&element)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return element, nil
}

func (auto *UIAutomation) GetFocusedElement() (*Element, error) {
	var element *Element

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

func (auto *UIAutomation) GetRootElementBuildCache(cacheRequest *CacheRequest) (*Element, error) {
	var element *Element

	hr, _, _ := syscall.SyscallN(
		auto.VTable().GetRootElementBuildCache,
		uintptr(unsafe.Pointer(auto)),
		uintptr(unsafe.Pointer(cacheRequest)),
		uintptr(unsafe.Pointer(&element)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return element, nil
}

func (auto *UIAutomation) ElementFromHandleBuildCache(hwnd syscall.Handle, cacheRequest *CacheRequest) (*Element, error) {
	var element *Element

	hr, _, _ := syscall.SyscallN(
		auto.VTable().ElementFromHandleBuildCache,
		uintptr(unsafe.Pointer(auto)),
		uintptr(hwnd),
		uintptr(unsafe.Pointer(cacheRequest)),
		uintptr(unsafe.Pointer(&element)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return element, nil
}

func (auto *UIAutomation) ElementFromPointBuildCache(point *Point, cacheRequest *CacheRequest) (*Element, error) {
	var element *Element

	hr, _, _ := syscall.SyscallN(
		auto.VTable().ElementFromPointBuildCache,
		uintptr(unsafe.Pointer(auto)),
		uintptr(unsafe.Pointer(point)),
		uintptr(unsafe.Pointer(cacheRequest)),
		uintptr(unsafe.Pointer(&element)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return element, nil
}

func (auto *UIAutomation) GetFocusedElementBuildCache(cacheRequest *CacheRequest) (*Element, error) {
	var element *Element

	hr, _, _ := syscall.SyscallN(
		auto.VTable().GetFocusedElementBuildCache,
		uintptr(unsafe.Pointer(auto)),
		uintptr(unsafe.Pointer(cacheRequest)),
		uintptr(unsafe.Pointer(&element)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return element, nil
}

func (auto *UIAutomation) CreateTreeWalker(condition *Condition) (*TreeWalker, error) {
	var walker *TreeWalker

	hr, _, _ := syscall.SyscallN(
		auto.VTable().CreateTreeWalker,
		uintptr(unsafe.Pointer(auto)),
		uintptr(unsafe.Pointer(condition)),
		uintptr(unsafe.Pointer(&walker)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return walker, nil
}

func (auto *UIAutomation) ControlViewWalker() (*TreeWalker, error) {
	var walker *TreeWalker

	hr, _, _ := syscall.SyscallN(
		auto.VTable().Get_ControlViewWalker,
		uintptr(unsafe.Pointer(auto)),
		uintptr(unsafe.Pointer(&walker)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return walker, nil
}

func (auto *UIAutomation) ContentViewWalker() (*TreeWalker, error) {
	var walker *TreeWalker

	hr, _, _ := syscall.SyscallN(
		auto.VTable().Get_ContentViewWalker,
		uintptr(unsafe.Pointer(auto)),
		uintptr(unsafe.Pointer(&walker)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return walker, nil
}

func (auto *UIAutomation) RawViewWalker() (*TreeWalker, error) {
	var walker *TreeWalker

	hr, _, _ := syscall.SyscallN(
		auto.VTable().Get_RawViewWalker,
		uintptr(unsafe.Pointer(auto)),
		uintptr(unsafe.Pointer(&walker)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return walker, nil
}

func (auto *UIAutomation) RawViewCondition() (*Condition, error) {
	var condition *Condition

	hr, _, _ := syscall.SyscallN(
		auto.VTable().Get_RawViewCondition,
		uintptr(unsafe.Pointer(auto)),
		uintptr(unsafe.Pointer(&condition)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return condition, nil
}

func (auto *UIAutomation) ControlViewCondition() (*Condition, error) {
	var condition *Condition

	hr, _, _ := syscall.SyscallN(
		auto.VTable().Get_ControlViewCondition,
		uintptr(unsafe.Pointer(auto)),
		uintptr(unsafe.Pointer(&condition)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return condition, nil
}

func (auto *UIAutomation) ContentViewCondition() (*Condition, error) {
	var condition *Condition

	hr, _, _ := syscall.SyscallN(
		auto.VTable().Get_ContentViewCondition,
		uintptr(unsafe.Pointer(auto)),
		uintptr(unsafe.Pointer(&condition)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return condition, nil
}

func (auto *UIAutomation) CreateCacheRequest() (*CacheRequest, error) {
	var cacheRequest *CacheRequest

	hr, _, _ := syscall.SyscallN(
		auto.VTable().CreateCacheRequest,
		uintptr(unsafe.Pointer(auto)),
		uintptr(unsafe.Pointer(&cacheRequest)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return cacheRequest, nil
}

func (auto *UIAutomation) CreateTrueCondition() (*Condition, error) {
	var condition *Condition

	hr, _, _ := syscall.SyscallN(
		auto.VTable().CreateTrueCondition,
		uintptr(unsafe.Pointer(auto)),
		uintptr(unsafe.Pointer(&condition)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return condition, nil
}

func (auto *UIAutomation) CreateFalseCondition() (*Condition, error) {
	var condition *Condition

	hr, _, _ := syscall.SyscallN(
		auto.VTable().CreateFalseCondition,
		uintptr(unsafe.Pointer(auto)),
		uintptr(unsafe.Pointer(&condition)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return condition, nil
}

func (auto *UIAutomation) CreatePropertyCondition(propertyId PropertyId, value *ole.VARIANT) (*Condition, error) {
	var condition *Condition

	hr, _, _ := syscall.SyscallN(
		auto.VTable().CreatePropertyCondition,
		uintptr(unsafe.Pointer(auto)),
		uintptr(propertyId),
		uintptr(unsafe.Pointer(value)),
		uintptr(unsafe.Pointer(&condition)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return condition, nil
}

func (auto *UIAutomation) CreatePropertyConditionEx(propertyId PropertyId, value *ole.VARIANT, flags PropertyConditionFlags) (*Condition, error) {
	var condition *Condition

	hr, _, _ := syscall.SyscallN(
		auto.VTable().CreatePropertyConditionEx,
		uintptr(unsafe.Pointer(auto)),
		uintptr(propertyId),
		uintptr(unsafe.Pointer(value)),
		uintptr(flags),
		uintptr(unsafe.Pointer(&condition)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return condition, nil
}

func (auto *UIAutomation) CreateAndCondition(condition1, condition2 *Condition) (*Condition, error) {
	var condition *Condition

	hr, _, _ := syscall.SyscallN(
		auto.VTable().CreateAndCondition,
		uintptr(unsafe.Pointer(auto)),
		uintptr(unsafe.Pointer(condition1)),
		uintptr(unsafe.Pointer(condition2)),
		uintptr(unsafe.Pointer(&condition)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return condition, nil
}

func (auto *UIAutomation) CreateAndConditionFromArray(conditions *ole.SafeArray) (*Condition, error) {
	var condition *Condition

	hr, _, _ := syscall.SyscallN(
		auto.VTable().CreateAndConditionFromArray,
		uintptr(unsafe.Pointer(auto)),
		uintptr(unsafe.Pointer(condition)),
		uintptr(unsafe.Pointer(&condition)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return condition, nil
}

func (auto *UIAutomation) CreateAndConditionFromNativeArray(conditions []int32) (*Condition, error) {
	var condition *Condition

	hr, _, _ := syscall.SyscallN(
		auto.VTable().CreateAndConditionFromNativeArray,
		uintptr(unsafe.Pointer(auto)),
		uintptr(unsafe.Pointer(&conditions[0])),
		uintptr(len(conditions)),
		uintptr(unsafe.Pointer(&condition)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return condition, nil
}

func (auto *UIAutomation) CreateOrCondition(condition1, condition2 *Condition) (*Condition, error) {
	var condition *Condition

	hr, _, _ := syscall.SyscallN(
		auto.VTable().CreateOrCondition,
		uintptr(unsafe.Pointer(auto)),
		uintptr(unsafe.Pointer(condition1)),
		uintptr(unsafe.Pointer(condition2)),
		uintptr(unsafe.Pointer(&condition)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return condition, nil
}

func (auto *UIAutomation) CreateOrConditionFromArray(conditions *ole.SafeArray) (*Condition, error) {
	var condition *Condition

	hr, _, _ := syscall.SyscallN(
		auto.VTable().CreateOrConditionFromArray,
		uintptr(unsafe.Pointer(auto)),
		uintptr(unsafe.Pointer(conditions)),
		uintptr(unsafe.Pointer(&condition)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return condition, nil
}

func (auto *UIAutomation) CreateOrConditionFromNativeArray(conditions []int32) (*Condition, error) {
	var condition *Condition

	hr, _, _ := syscall.SyscallN(
		auto.VTable().CreateOrConditionFromNativeArray,
		uintptr(unsafe.Pointer(auto)),
		uintptr(unsafe.Pointer(&conditions[0])),
		uintptr(len(conditions)),
		uintptr(unsafe.Pointer(&condition)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return condition, nil
}

func (auto *UIAutomation) CreateNotCondition(condition *Condition) (*Condition, error) {
	var notCondition *Condition

	hr, _, _ := syscall.SyscallN(
		auto.VTable().CreateNotCondition,
		uintptr(unsafe.Pointer(auto)),
		uintptr(unsafe.Pointer(condition)),
		uintptr(unsafe.Pointer(&notCondition)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return notCondition, nil
}

func (auto *UIAutomation) AddAutomationEventHandler(eventId EventId, element *Element, scope TreeScope, cacheRequest *CacheRequest, handler *EventHandler) error {
	hr, _, _ := syscall.SyscallN(
		auto.VTable().AddAutomationEventHandler,
		uintptr(unsafe.Pointer(auto)),
		uintptr(eventId),
		uintptr(unsafe.Pointer(element)),
		uintptr(scope),
		uintptr(unsafe.Pointer(cacheRequest)),
		uintptr(unsafe.Pointer(handler)),
	)

	if hr != 0 {
		return ole.NewError(hr)
	}

	return nil
}

func (auto *UIAutomation) RemoveAutomationEventHandler(eventId EventId, element *Element, handler *EventHandler) error {
	hr, _, _ := syscall.SyscallN(
		auto.VTable().RemoveAutomationEventHandler,
		uintptr(unsafe.Pointer(auto)),
		uintptr(eventId),
		uintptr(unsafe.Pointer(element)),
		uintptr(unsafe.Pointer(handler)),
	)

	if hr != 0 {
		return ole.NewError(hr)
	}

	return nil
}

func (auto *UIAutomation) AddPropertyChangedEventHandlerNativeArray(element *Element, scope TreeScope, cacheRequest *CacheRequest, handler *PropertyChangedEventHandler, propertyArray []PropertyId) error {
	hr, _, _ := syscall.SyscallN(
		auto.VTable().AddPropertyChangedEventHandlerNativeArray,
		uintptr(unsafe.Pointer(auto)),
		uintptr(unsafe.Pointer(element)),
		uintptr(scope),
		uintptr(unsafe.Pointer(cacheRequest)),
		uintptr(unsafe.Pointer(handler)),
		uintptr(unsafe.Pointer(&propertyArray[0])),
		uintptr(len(propertyArray)),
	)

	if hr != 0 {
		return ole.NewError(hr)
	}

	return nil
}

func (auto *UIAutomation) AddPropertyChangedEventHandler(element *Element, scope TreeScope, cacheRequest *CacheRequest, handler *PropertyChangedEventHandler, propertyArray *ole.SafeArray) error {
	hr, _, _ := syscall.SyscallN(
		auto.VTable().AddPropertyChangedEventHandlerNativeArray,
		uintptr(unsafe.Pointer(auto)),
		uintptr(unsafe.Pointer(element)),
		uintptr(scope),
		uintptr(unsafe.Pointer(cacheRequest)),
		uintptr(unsafe.Pointer(handler)),
		uintptr(unsafe.Pointer(propertyArray)),
	)

	if hr != 0 {
		return ole.NewError(hr)
	}

	return nil
}

func (auto *UIAutomation) RemovePropertyChangedEventHandler(element *Element, handler *PropertyChangedEventHandler) error {
	hr, _, _ := syscall.SyscallN(
		auto.VTable().RemovePropertyChangedEventHandler,
		uintptr(unsafe.Pointer(auto)),
		uintptr(unsafe.Pointer(element)),
		uintptr(unsafe.Pointer(handler)),
	)

	if hr != 0 {
		return ole.NewError(hr)
	}

	return nil
}

func (auto *UIAutomation) AddStructureChangedEventHandler(element *Element, scope TreeScope, cacheRequest *CacheRequest, handler *StructureChangedEventHandler) error {
	hr, _, _ := syscall.SyscallN(
		auto.VTable().AddStructureChangedEventHandler,
		uintptr(unsafe.Pointer(auto)),
		uintptr(unsafe.Pointer(element)),
		uintptr(scope),
		uintptr(unsafe.Pointer(cacheRequest)),
		uintptr(unsafe.Pointer(handler)),
	)

	if hr != 0 {
		return ole.NewError(hr)
	}

	return nil
}

func (auto *UIAutomation) RemoveStructureChangedEventHandler(element *Element, handler *StructureChangedEventHandler) error {
	hr, _, _ := syscall.SyscallN(
		auto.VTable().RemoveStructureChangedEventHandler,
		uintptr(unsafe.Pointer(auto)),
		uintptr(unsafe.Pointer(element)),
		uintptr(unsafe.Pointer(handler)),
	)

	if hr != 0 {
		return ole.NewError(hr)
	}

	return nil
}

func (auto *UIAutomation) AddFocusChangedEventHandler(handler *FocusChangedEventHandler) error {
	hr, _, _ := syscall.SyscallN(
		auto.VTable().AddFocusChangedEventHandler,
		uintptr(unsafe.Pointer(auto)),
		uintptr(unsafe.Pointer(handler)),
	)

	if hr != 0 {
		return ole.NewError(hr)
	}

	return nil
}

func (auto *UIAutomation) RemoveFocusChangedEventHandler(handler *FocusChangedEventHandler) error {
	hr, _, _ := syscall.SyscallN(
		auto.VTable().RemoveFocusChangedEventHandler,
		uintptr(unsafe.Pointer(auto)),
		uintptr(unsafe.Pointer(handler)),
	)

	if hr != 0 {
		return ole.NewError(hr)
	}

	return nil
}

func (auto *UIAutomation) RemoveAllEventHandlers() error {
	hr, _, _ := syscall.SyscallN(
		auto.VTable().RemoveAllEventHandlers,
		uintptr(unsafe.Pointer(auto)),
	)

	if hr != 0 {
		return ole.NewError(hr)
	}

	return nil
}

func (auto *UIAutomation) IntNativeArrayToSafeArray(array []int32) (*ole.SafeArray, error) {
	var safeArray *ole.SafeArray

	hr, _, _ := syscall.SyscallN(
		auto.VTable().IntNativeArrayToSafeArray,
		uintptr(unsafe.Pointer(auto)),
		uintptr(unsafe.Pointer(&array[0])),
		uintptr(len(array)),
		uintptr(unsafe.Pointer(&safeArray)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return safeArray, nil
}

func (auto *UIAutomation) IntSafeArrayToNativeArray(safeArray *ole.SafeArray) ([]int32, error) {
	var array *int32
	var arrayCount int32

	hr, _, _ := syscall.SyscallN(
		auto.VTable().IntSafeArrayToNativeArray,
		uintptr(unsafe.Pointer(auto)),
		uintptr(unsafe.Pointer(safeArray)),
		uintptr(unsafe.Pointer(&array)),
		uintptr(unsafe.Pointer(&arrayCount)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	slice := (*[1 << 30]int32)(unsafe.Pointer(array))[:arrayCount:arrayCount]
	result := make([]int32, arrayCount)
	copy(result, slice)

	ole.CoTaskMemFree(uintptr(unsafe.Pointer(array)))

	return result, nil
}

func (auto *UIAutomation) RectToVariant(rect Rect) (*ole.VARIANT, error) {
	var variant *ole.VARIANT

	hr, _, _ := syscall.SyscallN(
		auto.VTable().RectToVariant,
		uintptr(unsafe.Pointer(auto)),
		uintptr(unsafe.Pointer(&rect)),
		uintptr(unsafe.Pointer(variant)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return variant, nil
}

func (auto *UIAutomation) VariantToRect(variant *ole.VARIANT) (Rect, error) {
	var rect Rect

	hr, _, _ := syscall.SyscallN(
		auto.VTable().VariantToRect,
		uintptr(unsafe.Pointer(auto)),
		uintptr(unsafe.Pointer(variant)),
		uintptr(unsafe.Pointer(&rect)),
	)

	if hr != 0 {
		return Rect{}, ole.NewError(hr)
	}

	return rect, nil
}

func (auto *UIAutomation) SafeArrayToRectNativeArray(safeArray *ole.SafeArray) ([]Rect, error) {
	var rectArray *Rect
	var rectArrayCount int32

	hr, _, _ := syscall.SyscallN(
		auto.VTable().SafeArrayToRectNativeArray,
		uintptr(unsafe.Pointer(auto)),
		uintptr(unsafe.Pointer(safeArray)),
		uintptr(unsafe.Pointer(&rectArray)),
		uintptr(unsafe.Pointer(&rectArrayCount)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	slice := (*[1 << 30]Rect)(unsafe.Pointer(rectArray))[:rectArrayCount:rectArrayCount]
	result := make([]Rect, rectArrayCount)
	copy(result, slice)

	ole.CoTaskMemFree(uintptr(unsafe.Pointer(rectArray)))

	return result, nil
}

func (auto *UIAutomation) CreateProxyFactoryEntry(factory *ProxyFactory) (*ProxyFactoryEntry, error) {
	var entry *ProxyFactoryEntry

	hr, _, _ := syscall.SyscallN(
		auto.VTable().CreateProxyFactoryEntry,
		uintptr(unsafe.Pointer(auto)),
		uintptr(unsafe.Pointer(factory)),
		uintptr(unsafe.Pointer(&entry)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return entry, nil
}

func (auto *UIAutomation) ProxyFactoryMapping() (*ProxyFactoryMapping, error) {
	var mapping *ProxyFactoryMapping

	hr, _, _ := syscall.SyscallN(
		auto.VTable().Get_ProxyFactoryMapping,
		uintptr(unsafe.Pointer(auto)),
		uintptr(unsafe.Pointer(&mapping)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return mapping, nil
}

func (auto *UIAutomation) GetPropertyProgrammaticName(propertyId PropertyId) (string, error) {
	var name *uint16

	hr, _, _ := syscall.SyscallN(
		auto.VTable().GetPropertyProgrammaticName,
		uintptr(unsafe.Pointer(auto)),
		uintptr(propertyId),
		uintptr(unsafe.Pointer(&name)),
	)

	if hr != 0 {
		return "", ole.NewError(hr)
	}

	str := ole.BstrToString(name)
	ole.CoTaskMemFree(uintptr(unsafe.Pointer(name)))

	return str, nil
}
func (auto *UIAutomation) GetPatternProgrammaticName(patternId PatternId) (string, error) {
	var name *uint16

	hr, _, _ := syscall.SyscallN(
		auto.VTable().GetPatternProgrammaticName,
		uintptr(unsafe.Pointer(auto)),
		uintptr(patternId),
		uintptr(unsafe.Pointer(&name)),
	)

	if hr != 0 {
		return "", ole.NewError(hr)
	}

	str := ole.BstrToString(name)
	ole.CoTaskMemFree(uintptr(unsafe.Pointer(name)))

	return str, nil
}

func (auto *UIAutomation) PollForPotentialSupportedPatterns(element *Element) ([]PatternId, error) {
	var patternIds *ole.SafeArray
	var patternNames *ole.SafeArray

	hr, _, _ := syscall.SyscallN(
		auto.VTable().PollForPotentialSupportedPatterns,
		uintptr(unsafe.Pointer(auto)),
		uintptr(unsafe.Pointer(element)),
		uintptr(unsafe.Pointer(&patternIds)),
		uintptr(unsafe.Pointer(&patternNames)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	array, err := auto.IntSafeArrayToNativeArray(patternIds)

	if err != nil {
		return nil, err
	}

	patterns := make([]PatternId, len(array))

	for i, id := range array {
		patterns[i] = PatternId(id)
	}

	return patterns, nil
}

func (auto *UIAutomation) PollForPotentialSupportedProperties(element *Element) ([]PropertyId, error) {
	var propertyIds *ole.SafeArray
	var propertyNames *ole.SafeArray

	hr, _, _ := syscall.SyscallN(
		auto.VTable().PollForPotentialSupportedProperties,
		uintptr(unsafe.Pointer(auto)),
		uintptr(unsafe.Pointer(element)),
		uintptr(unsafe.Pointer(&propertyIds)),
		uintptr(unsafe.Pointer(&propertyNames)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	array, err := auto.IntSafeArrayToNativeArray(propertyIds)

	if err != nil {
		return nil, err
	}

	properties := make([]PropertyId, len(array))

	for i, id := range array {
		properties[i] = PropertyId(id)
	}

	return properties, nil
}

func (auto *UIAutomation) CheckNotSupported(value *ole.VARIANT) (bool, error) {
	var isNotSupported bool

	hr, _, _ := syscall.SyscallN(
		auto.VTable().CheckNotSupported,
		uintptr(unsafe.Pointer(auto)),
		uintptr(unsafe.Pointer(value)),
		uintptr(unsafe.Pointer(&isNotSupported)),
	)

	if hr != 0 {
		return false, ole.NewError(hr)
	}

	return isNotSupported, nil
}

func (auto *UIAutomation) ReservedNotSupportedValue() (*ole.IUnknown, error) {
	var value *ole.IUnknown

	hr, _, _ := syscall.SyscallN(
		auto.VTable().Get_ReservedNotSupportedValue,
		uintptr(unsafe.Pointer(auto)),
		uintptr(unsafe.Pointer(&value)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return value, nil
}

func (auto *UIAutomation) ReservedMixedAttributeValue() (*ole.IUnknown, error) {
	var value *ole.IUnknown

	hr, _, _ := syscall.SyscallN(
		auto.VTable().Get_ReservedMixedAttributeValue,
		uintptr(unsafe.Pointer(auto)),
		uintptr(unsafe.Pointer(&value)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return value, nil
}

func (auto *UIAutomation) ElementFromIAccessible(accessible *Accessible, childId int32) (*Element, error) {
	var element *Element

	hr, _, _ := syscall.SyscallN(
		auto.VTable().ElementFromIAccessible,
		uintptr(unsafe.Pointer(auto)),
		uintptr(unsafe.Pointer(accessible)),
		uintptr(childId),
		uintptr(unsafe.Pointer(&element)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return element, nil
}

func (auto *UIAutomation) ElementFromIAccessibleBuildCache(accessible *Accessible, childId int32, cacheRequest *CacheRequest) (*Element, error) {
	var element *Element

	hr, _, _ := syscall.SyscallN(
		auto.VTable().ElementFromIAccessibleBuildCache,
		uintptr(unsafe.Pointer(auto)),
		uintptr(unsafe.Pointer(accessible)),
		uintptr(childId),
		uintptr(unsafe.Pointer(cacheRequest)),
		uintptr(unsafe.Pointer(&element)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return element, nil
}
