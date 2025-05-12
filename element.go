package uiautomation

import (
	"syscall"
	"unsafe"

	"github.com/go-ole/go-ole"
)

type Element struct {
	ole.IUnknown
}

type ElementVtbl struct {
	ole.IUnknownVtbl
	SetFocus                        uintptr
	GetRuntimeId                    uintptr
	FindFirst                       uintptr
	FindAll                         uintptr
	FindFirstBuildCache             uintptr
	FindAllBuildCache               uintptr
	BuildUpdatedCache               uintptr
	GetCurrentPropertyValue         uintptr
	GetCurrentPropertyValueEx       uintptr
	GetCachedPropertyValue          uintptr
	GetCachedPropertyValueEx        uintptr
	GetCurrentPatternAs             uintptr
	GetCachedPatternAs              uintptr
	GetCurrentPattern               uintptr
	GetCachedPattern                uintptr
	GetCachedParent                 uintptr
	GetCachedChildren               uintptr
	Get_CurrentProcessId            uintptr
	Get_CurrentControlType          uintptr
	Get_CurrentLocalizedControlType uintptr
	Get_CurrentName                 uintptr
	Get_CurrentAcceleratorKey       uintptr
	Get_CurrentAccessKey            uintptr
	Get_CurrentHasKeyboardFocus     uintptr
	Get_CurrentIsKeyboardFocusable  uintptr
	Get_CurrentIsEnabled            uintptr
	Get_CurrentAutomationId         uintptr
	Get_CurrentClassName            uintptr
	Get_CurrentHelpText             uintptr
	Get_CurrentCulture              uintptr
	Get_CurrentIsControlElement     uintptr
	Get_CurrentIsContentElement     uintptr
	Get_CurrentIsPassword           uintptr
	Get_CurrentNativeWindowHandle   uintptr
	Get_CurrentItemType             uintptr
	Get_CurrentIsOffscreen          uintptr
	Get_CurrentOrientation          uintptr
	Get_CurrentFrameworkId          uintptr
	Get_CurrentIsRequiredForForm    uintptr
	Get_CurrentItemStatus           uintptr
	Get_CurrentBoundingRectangle    uintptr
	Get_CurrentLabeledBy            uintptr
	Get_CurrentAriaRole             uintptr
	Get_CurrentAriaProperties       uintptr
	Get_CurrentIsDataValidForForm   uintptr
	Get_CurrentControllerFor        uintptr
	Get_CurrentDescribedBy          uintptr
	Get_CurrentFlowsTo              uintptr
	Get_CurrentProviderDescription  uintptr
	Get_CachedProcessId             uintptr
	Get_CachedControlType           uintptr
	Get_CachedLocalizedControlType  uintptr
	Get_CachedName                  uintptr
	Get_CachedAcceleratorKey        uintptr
	Get_CachedAccessKey             uintptr
	Get_CachedHasKeyboardFocus      uintptr
	Get_CachedIsKeyboardFocusable   uintptr
	Get_CachedIsEnabled             uintptr
	Get_CachedAutomationId          uintptr
	Get_CachedClassName             uintptr
	Get_CachedHelpText              uintptr
	Get_CachedCulture               uintptr
	Get_CachedIsControlElement      uintptr
	Get_CachedIsContentElement      uintptr
	Get_CachedIsPassword            uintptr
	Get_CachedNativeWindowHandle    uintptr
	Get_CachedItemType              uintptr
	Get_CachedIsOffscreen           uintptr
	Get_CachedOrientation           uintptr
	Get_CachedFrameworkId           uintptr
	Get_CachedIsRequiredForForm     uintptr
	Get_CachedItemStatus            uintptr
	Get_CachedBoundingRectangle     uintptr
	Get_CachedLabeledBy             uintptr
	Get_CachedAriaRole              uintptr
	Get_CachedAriaProperties        uintptr
	Get_CachedIsDataValidForForm    uintptr
	Get_CachedControllerFor         uintptr
	Get_CachedDescribedBy           uintptr
	Get_CachedFlowsTo               uintptr
	Get_CachedProviderDescription   uintptr
	GetClickablePoint               uintptr
}

func (elem *Element) VTable() *ElementVtbl {
	return (*ElementVtbl)(unsafe.Pointer(elem.RawVTable))
}

func (elem *Element) SetFocus() error {
	hr, _, _ := syscall.SyscallN(
		elem.VTable().SetFocus,
		uintptr(unsafe.Pointer(elem)),
	)

	if hr != 0 {
		return ole.NewError(hr)
	}

	return nil
}

func (elem *Element) GetRuntimeId() (*ole.SafeArray, error) {
	var runtimeId *ole.SafeArray

	hr, _, _ := syscall.SyscallN(
		elem.VTable().GetRuntimeId,
		uintptr(unsafe.Pointer(elem)),
		uintptr(unsafe.Pointer(&runtimeId)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return runtimeId, nil
}

func (elem *Element) FindFirst(scope TreeScope, condition *Condition) (*Element, error) {
	var found *Element

	hr, _, _ := syscall.SyscallN(
		elem.VTable().FindFirst,
		uintptr(unsafe.Pointer(elem)),
		uintptr(scope),
		uintptr(unsafe.Pointer(condition)),
		uintptr(unsafe.Pointer(&found)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return found, nil
}

func (elem *Element) FindAll(scope TreeScope, condition *Condition) (*ElementArray, error) {
	var found *ElementArray

	hr, _, _ := syscall.SyscallN(
		elem.VTable().FindAll,
		uintptr(unsafe.Pointer(elem)),
		uintptr(scope),
		uintptr(unsafe.Pointer(condition)),
		uintptr(unsafe.Pointer(&found)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return found, nil
}

func (elem *Element) FindFirstBuildCache(scope TreeScope, condition *Condition, cacheRequest *CacheRequest) (*Element, error) {
	var found *Element

	hr, _, _ := syscall.SyscallN(
		elem.VTable().FindFirstBuildCache,
		uintptr(unsafe.Pointer(elem)),
		uintptr(scope),
		uintptr(unsafe.Pointer(condition)),
		uintptr(unsafe.Pointer(cacheRequest)),
		uintptr(unsafe.Pointer(&found)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return found, nil
}

func (elem *Element) FindAllBuildCache(scope TreeScope, condition *Condition, cacheRequest *CacheRequest) (*ElementArray, error) {
	var found *ElementArray

	hr, _, _ := syscall.SyscallN(
		elem.VTable().FindAllBuildCache,
		uintptr(unsafe.Pointer(elem)),
		uintptr(scope),
		uintptr(unsafe.Pointer(condition)),
		uintptr(unsafe.Pointer(cacheRequest)),
		uintptr(unsafe.Pointer(&found)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return found, nil
}

func (elem *Element) BuildUpdatedCache(cacheRequest *CacheRequest) (*Element, error) {
	var updated *Element

	hr, _, _ := syscall.SyscallN(
		elem.VTable().BuildUpdatedCache,
		uintptr(unsafe.Pointer(elem)),
		uintptr(unsafe.Pointer(cacheRequest)),
		uintptr(unsafe.Pointer(&updated)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return updated, nil
}

func (elem *Element) GetCurrentPropertyValue(propertyId PropertyId) (*ole.VARIANT, error) {
	value := &ole.VARIANT{}

	ole.VariantInit(value)

	hr, _, _ := syscall.SyscallN(
		elem.VTable().GetCurrentPropertyValue,
		uintptr(unsafe.Pointer(elem)),
		uintptr(propertyId),
		uintptr(unsafe.Pointer(value)),
	)

	if hr != 0 {
		ole.VariantClear(value)
		return nil, ole.NewError(hr)
	}

	return value, nil
}

func (elem *Element) GetCurrentPropertyValueEx(propertyId PropertyId, ignoreDefaultValue bool) (*ole.VARIANT, error) {
	value := &ole.VARIANT{}

	ole.VariantInit(value)

	hr, _, _ := syscall.SyscallN(
		elem.VTable().GetCurrentPropertyValueEx,
		uintptr(unsafe.Pointer(elem)),
		uintptr(propertyId),
		uintptr(unsafe.Pointer(&ignoreDefaultValue)),
		uintptr(unsafe.Pointer(value)),
	)

	if hr != 0 {
		ole.VariantClear(value)
		return nil, ole.NewError(hr)
	}

	return value, nil
}

func (elem *Element) GetCachedPropertyValue(propertyId PropertyId) (*ole.VARIANT, error) {
	value := &ole.VARIANT{}

	ole.VariantInit(value)

	hr, _, _ := syscall.SyscallN(
		elem.VTable().GetCachedPropertyValue,
		uintptr(unsafe.Pointer(elem)),
		uintptr(propertyId),
		uintptr(unsafe.Pointer(value)),
	)

	if hr != 0 {
		ole.VariantClear(value)
		return nil, ole.NewError(hr)
	}

	return value, nil
}

func (elem *Element) GetCachedPropertyValueEx(propertyId PropertyId, ignoreDefaultValue bool) (*ole.VARIANT, error) {
	value := &ole.VARIANT{}

	ole.VariantInit(value)

	hr, _, _ := syscall.SyscallN(
		elem.VTable().GetCachedPropertyValueEx,
		uintptr(unsafe.Pointer(elem)),
		uintptr(propertyId),
		uintptr(unsafe.Pointer(&ignoreDefaultValue)),
		uintptr(unsafe.Pointer(value)),
	)

	if hr != 0 {
		ole.VariantClear(value)
		return nil, ole.NewError(hr)
	}

	return value, nil
}

func (elem *Element) GetCurrentPatternAs(patternId PatternId, riid *ole.GUID) (*ole.IUnknown, error) {
	var patternObject *ole.IUnknown

	hr, _, _ := syscall.SyscallN(
		elem.VTable().GetCurrentPatternAs,
		uintptr(unsafe.Pointer(elem)),
		uintptr(patternId),
		uintptr(unsafe.Pointer(riid)),
		uintptr(unsafe.Pointer(&patternObject)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return patternObject, nil
}

func (elem *Element) GetCachedPatternAs(patternId PatternId, riid *ole.GUID) (*ole.IUnknown, error) {
	var patternObject *ole.IUnknown

	hr, _, _ := syscall.SyscallN(
		elem.VTable().GetCachedPatternAs,
		uintptr(unsafe.Pointer(elem)),
		uintptr(patternId),
		uintptr(unsafe.Pointer(riid)),
		uintptr(unsafe.Pointer(&patternObject)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return patternObject, nil
}

func (elem *Element) GetCurrentPattern(patternId PatternId) (*ole.IUnknown, error) {
	var patternObject *ole.IUnknown

	hr, _, _ := syscall.SyscallN(
		elem.VTable().GetCurrentPattern,
		uintptr(unsafe.Pointer(elem)),
		uintptr(patternId),
		uintptr(unsafe.Pointer(&patternObject)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return patternObject, nil
}

func (elem *Element) GetCachedPattern(patternId PatternId) (*ole.IUnknown, error) {
	var patternObject *ole.IUnknown

	hr, _, _ := syscall.SyscallN(
		elem.VTable().GetCachedPattern,
		uintptr(unsafe.Pointer(elem)),
		uintptr(patternId),
		uintptr(unsafe.Pointer(&patternObject)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return patternObject, nil
}

func (elem *Element) GetCachedParent() (*Element, error) {
	var parent *Element

	hr, _, _ := syscall.SyscallN(
		elem.VTable().GetCachedParent,
		uintptr(unsafe.Pointer(elem)),
		uintptr(unsafe.Pointer(&parent)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return parent, nil
}

func (elem *Element) GetCachedChildren() (*ElementArray, error) {
	var children *ElementArray

	hr, _, _ := syscall.SyscallN(
		elem.VTable().GetCachedChildren,
		uintptr(unsafe.Pointer(elem)),
		uintptr(unsafe.Pointer(&children)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return children, nil
}

func (elem *Element) CurrentProcessId() (int32, error) {
	var retVal int32

	hr, _, _ := syscall.SyscallN(
		elem.VTable().Get_CurrentProcessId,
		uintptr(unsafe.Pointer(elem)),
		uintptr(unsafe.Pointer(&retVal)),
	)

	if hr != 0 {
		return 0, ole.NewError(hr)
	}

	return retVal, nil
}

func (elem *Element) CurrentControlType() (ControlTypeId, error) {
	var retVal ControlTypeId

	hr, _, _ := syscall.SyscallN(
		elem.VTable().Get_CurrentControlType,
		uintptr(unsafe.Pointer(elem)),
		uintptr(unsafe.Pointer(&retVal)),
	)

	if hr != 0 {
		return 0, ole.NewError(hr)
	}

	return retVal, nil
}

func (elem *Element) CurrentLocalizedControlType() (string, error) {
	var retVal *uint16

	hr, _, _ := syscall.SyscallN(
		elem.VTable().Get_CurrentLocalizedControlType,
		uintptr(unsafe.Pointer(elem)),
		uintptr(unsafe.Pointer(&retVal)),
	)

	if hr != 0 {
		return "", ole.NewError(hr)
	}

	return ole.BstrToString(retVal), nil
}

func (elem *Element) CurrentName() (string, error) {
	var retVal *uint16

	hr, _, _ := syscall.SyscallN(
		elem.VTable().Get_CurrentName,
		uintptr(unsafe.Pointer(elem)),
		uintptr(unsafe.Pointer(&retVal)),
	)

	if hr != 0 {
		return "", ole.NewError(hr)
	}

	return ole.BstrToString(retVal), nil
}

func (elem *Element) CurrentAcceleratorKey() (string, error) {
	var retVal *uint16

	hr, _, _ := syscall.SyscallN(
		elem.VTable().Get_CurrentAcceleratorKey,
		uintptr(unsafe.Pointer(elem)),
		uintptr(unsafe.Pointer(&retVal)),
	)

	if hr != 0 {
		return "", ole.NewError(hr)
	}

	return ole.BstrToString(retVal), nil
}

func (elem *Element) CurrentAccessKey() (string, error) {
	var retVal *uint16

	hr, _, _ := syscall.SyscallN(
		elem.VTable().Get_CurrentAccessKey,
		uintptr(unsafe.Pointer(elem)),
		uintptr(unsafe.Pointer(&retVal)),
	)

	if hr != 0 {
		return "", ole.NewError(hr)
	}

	return ole.BstrToString(retVal), nil
}

func (elem *Element) CurrentHasKeyboardFocus() (bool, error) {
	var retVal bool

	hr, _, _ := syscall.SyscallN(
		elem.VTable().Get_CurrentHasKeyboardFocus,
		uintptr(unsafe.Pointer(elem)),
		uintptr(unsafe.Pointer(&retVal)),
	)

	if hr != 0 {
		return false, ole.NewError(hr)
	}

	return retVal, nil
}

func (elem *Element) CurrentIsKeyboardFocusable() (bool, error) {
	var retVal bool

	hr, _, _ := syscall.SyscallN(
		elem.VTable().Get_CurrentIsKeyboardFocusable,
		uintptr(unsafe.Pointer(elem)),
		uintptr(unsafe.Pointer(&retVal)),
	)

	if hr != 0 {
		return false, ole.NewError(hr)
	}

	return retVal, nil
}

func (elem *Element) CurrentIsEnabled() (bool, error) {
	var retVal bool

	hr, _, _ := syscall.SyscallN(
		elem.VTable().Get_CurrentIsEnabled,
		uintptr(unsafe.Pointer(elem)),
		uintptr(unsafe.Pointer(&retVal)),
	)

	if hr != 0 {
		return false, ole.NewError(hr)
	}

	return retVal, nil
}

func (elem *Element) CurrentAutomationId() (string, error) {
	var retVal *uint16

	hr, _, _ := syscall.SyscallN(
		elem.VTable().Get_CurrentAutomationId,
		uintptr(unsafe.Pointer(elem)),
		uintptr(unsafe.Pointer(&retVal)),
	)

	if hr != 0 {
		return "", ole.NewError(hr)
	}

	return ole.BstrToString(retVal), nil
}

func (elem *Element) CurrentClassName() (string, error) {
	var retVal *uint16

	hr, _, _ := syscall.SyscallN(
		elem.VTable().Get_CurrentClassName,
		uintptr(unsafe.Pointer(elem)),
		uintptr(unsafe.Pointer(&retVal)),
	)

	if hr != 0 {
		return "", ole.NewError(hr)
	}

	return ole.BstrToString(retVal), nil
}

func (elem *Element) CurrentHelpText() (string, error) {
	var retVal *uint16

	hr, _, _ := syscall.SyscallN(
		elem.VTable().Get_CurrentHelpText,
		uintptr(unsafe.Pointer(elem)),
		uintptr(unsafe.Pointer(&retVal)),
	)

	if hr != 0 {
		return "", ole.NewError(hr)
	}

	return ole.BstrToString(retVal), nil
}

func (elem *Element) CurrentCulture() (int32, error) {
	var retVal int32

	hr, _, _ := syscall.SyscallN(
		elem.VTable().Get_CurrentCulture,
		uintptr(unsafe.Pointer(elem)),
		uintptr(unsafe.Pointer(&retVal)),
	)

	if hr != 0 {
		return 0, ole.NewError(hr)
	}

	return retVal, nil
}

func (elem *Element) CurrentIsControlElement() (bool, error) {
	var retVal bool

	hr, _, _ := syscall.SyscallN(
		elem.VTable().Get_CurrentIsControlElement,
		uintptr(unsafe.Pointer(elem)),
		uintptr(unsafe.Pointer(&retVal)),
	)

	if hr != 0 {
		return false, ole.NewError(hr)
	}

	return retVal, nil
}

func (elem *Element) CurrentIsContentElement() (bool, error) {
	var retVal bool

	hr, _, _ := syscall.SyscallN(
		elem.VTable().Get_CurrentIsContentElement,
		uintptr(unsafe.Pointer(elem)),
		uintptr(unsafe.Pointer(&retVal)),
	)

	if hr != 0 {
		return false, ole.NewError(hr)
	}

	return retVal, nil
}

func (elem *Element) CurrentIsPassword() (bool, error) {
	var retVal bool

	hr, _, _ := syscall.SyscallN(
		elem.VTable().Get_CurrentIsPassword,
		uintptr(unsafe.Pointer(elem)),
		uintptr(unsafe.Pointer(&retVal)),
	)

	if hr != 0 {
		return false, ole.NewError(hr)
	}

	return retVal, nil
}

func (elem *Element) CurrentNativeWindowHandle() (syscall.Handle, error) {
	var retVal syscall.Handle

	hr, _, _ := syscall.SyscallN(
		elem.VTable().Get_CurrentNativeWindowHandle,
		uintptr(unsafe.Pointer(elem)),
		uintptr(unsafe.Pointer(&retVal)),
	)

	if hr != 0 {
		return 0, ole.NewError(hr)
	}

	return retVal, nil
}

func (elem *Element) CurrentItemType() (string, error) {
	var retVal *uint16

	hr, _, _ := syscall.SyscallN(
		elem.VTable().Get_CurrentItemType,
		uintptr(unsafe.Pointer(elem)),
		uintptr(unsafe.Pointer(&retVal)),
	)

	if hr != 0 {
		return "", ole.NewError(hr)
	}

	return ole.BstrToString(retVal), nil
}

func (elem *Element) CurrentIsOffscreen() (bool, error) {
	var retVal bool

	hr, _, _ := syscall.SyscallN(
		elem.VTable().Get_CurrentIsOffscreen,
		uintptr(unsafe.Pointer(elem)),
		uintptr(unsafe.Pointer(&retVal)),
	)

	if hr != 0 {
		return false, ole.NewError(hr)
	}

	return retVal, nil
}

func (elem *Element) CurrentOrientation() (OrientationType, error) {
	var retVal OrientationType

	hr, _, _ := syscall.SyscallN(
		elem.VTable().Get_CurrentOrientation,
		uintptr(unsafe.Pointer(elem)),
		uintptr(unsafe.Pointer(&retVal)),
	)

	if hr != 0 {
		return 0, ole.NewError(hr)
	}

	return retVal, nil
}

func (elem *Element) CurrentFrameworkId() (string, error) {
	var retVal *uint16

	hr, _, _ := syscall.SyscallN(
		elem.VTable().Get_CurrentFrameworkId,
		uintptr(unsafe.Pointer(elem)),
		uintptr(unsafe.Pointer(&retVal)),
	)

	if hr != 0 {
		return "", ole.NewError(hr)
	}

	return ole.BstrToString(retVal), nil
}

func (elem *Element) CurrentIsRequiredForForm() (bool, error) {
	var retVal bool

	hr, _, _ := syscall.SyscallN(
		elem.VTable().Get_CurrentIsRequiredForForm,
		uintptr(unsafe.Pointer(elem)),
		uintptr(unsafe.Pointer(&retVal)),
	)

	if hr != 0 {
		return false, ole.NewError(hr)
	}

	return retVal, nil
}

func (elem *Element) CurrentItemStatus() (string, error) {
	var retVal *uint16

	hr, _, _ := syscall.SyscallN(
		elem.VTable().Get_CurrentItemStatus,
		uintptr(unsafe.Pointer(elem)),
		uintptr(unsafe.Pointer(&retVal)),
	)

	if hr != 0 {
		return "", ole.NewError(hr)
	}

	return ole.BstrToString(retVal), nil
}

func (elem *Element) CurrentBoundingRectangle() (Rect, error) {
	var retVal Rect

	hr, _, _ := syscall.SyscallN(
		elem.VTable().Get_CurrentBoundingRectangle,
		uintptr(unsafe.Pointer(elem)),
		uintptr(unsafe.Pointer(&retVal)),
	)

	if hr != 0 {
		return Rect{}, ole.NewError(hr)
	}

	return retVal, nil
}

func (elem *Element) CurrentLabeledBy() (*Element, error) {
	var retVal *Element

	hr, _, _ := syscall.SyscallN(
		elem.VTable().Get_CurrentLabeledBy,
		uintptr(unsafe.Pointer(elem)),
		uintptr(unsafe.Pointer(&retVal)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return retVal, nil
}

func (elem *Element) CurrentAriaRole() (string, error) {
	var retVal *uint16

	hr, _, _ := syscall.SyscallN(
		elem.VTable().Get_CurrentAriaRole,
		uintptr(unsafe.Pointer(elem)),
		uintptr(unsafe.Pointer(&retVal)),
	)

	if hr != 0 {
		return "", ole.NewError(hr)
	}

	return ole.BstrToString(retVal), nil
}

func (elem *Element) CurrentAriaProperties() (string, error) {
	var retVal *uint16

	hr, _, _ := syscall.SyscallN(
		elem.VTable().Get_CurrentAriaProperties,
		uintptr(unsafe.Pointer(elem)),
		uintptr(unsafe.Pointer(&retVal)),
	)

	if hr != 0 {
		return "", ole.NewError(hr)
	}

	return ole.BstrToString(retVal), nil
}

func (elem *Element) CurrentIsDataValidForForm() (bool, error) {
	var retVal bool

	hr, _, _ := syscall.SyscallN(
		elem.VTable().Get_CurrentIsDataValidForForm,
		uintptr(unsafe.Pointer(elem)),
		uintptr(unsafe.Pointer(&retVal)),
	)

	if hr != 0 {
		return false, ole.NewError(hr)
	}

	return retVal, nil
}

func (elem *Element) CurrentControllerFor() (*ElementArray, error) {
	var retVal *ElementArray

	hr, _, _ := syscall.SyscallN(
		elem.VTable().Get_CurrentControllerFor,
		uintptr(unsafe.Pointer(elem)),
		uintptr(unsafe.Pointer(&retVal)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return retVal, nil
}

func (elem *Element) CurrentDescribedBy() (*ElementArray, error) {
	var retVal *ElementArray

	hr, _, _ := syscall.SyscallN(
		elem.VTable().Get_CurrentDescribedBy,
		uintptr(unsafe.Pointer(elem)),
		uintptr(unsafe.Pointer(&retVal)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return retVal, nil
}

func (elem *Element) CurrentFlowsTo() (*ElementArray, error) {
	var retVal *ElementArray

	hr, _, _ := syscall.SyscallN(
		elem.VTable().Get_CurrentFlowsTo,
		uintptr(unsafe.Pointer(elem)),
		uintptr(unsafe.Pointer(&retVal)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return retVal, nil
}

func (elem *Element) CurrentProviderDescription() (string, error) {
	var retVal *uint16

	hr, _, _ := syscall.SyscallN(
		elem.VTable().Get_CurrentProviderDescription,
		uintptr(unsafe.Pointer(elem)),
		uintptr(unsafe.Pointer(&retVal)),
	)

	if hr != 0 {
		return "", ole.NewError(hr)
	}

	return ole.BstrToString(retVal), nil
}

func (elem *Element) CachedProcessId() (int32, error) {
	var retVal int32

	hr, _, _ := syscall.SyscallN(
		elem.VTable().Get_CachedProcessId,
		uintptr(unsafe.Pointer(elem)),
		uintptr(unsafe.Pointer(&retVal)),
	)

	if hr != 0 {
		return 0, ole.NewError(hr)
	}

	return retVal, nil
}

func (elem *Element) CachedControlType() (ControlTypeId, error) {
	var retVal ControlTypeId

	hr, _, _ := syscall.SyscallN(
		elem.VTable().Get_CachedControlType,
		uintptr(unsafe.Pointer(elem)),
		uintptr(unsafe.Pointer(&retVal)),
	)

	if hr != 0 {
		return 0, ole.NewError(hr)
	}

	return retVal, nil
}

func (elem *Element) CachedLocalizedControlType() (string, error) {
	var retVal *uint16

	hr, _, _ := syscall.SyscallN(
		elem.VTable().Get_CachedLocalizedControlType,
		uintptr(unsafe.Pointer(elem)),
		uintptr(unsafe.Pointer(&retVal)),
	)

	if hr != 0 {
		return "", ole.NewError(hr)
	}

	return ole.BstrToString(retVal), nil
}

func (elem *Element) CachedName() (string, error) {
	var retVal *uint16

	hr, _, _ := syscall.SyscallN(
		elem.VTable().Get_CachedName,
		uintptr(unsafe.Pointer(elem)),
		uintptr(unsafe.Pointer(&retVal)),
	)

	if hr != 0 {
		return "", ole.NewError(hr)
	}

	return ole.BstrToString(retVal), nil
}

func (elem *Element) CachedAcceleratorKey() (string, error) {
	var retVal *uint16

	hr, _, _ := syscall.SyscallN(
		elem.VTable().Get_CachedAcceleratorKey,
		uintptr(unsafe.Pointer(elem)),
		uintptr(unsafe.Pointer(&retVal)),
	)

	if hr != 0 {
		return "", ole.NewError(hr)
	}

	return ole.BstrToString(retVal), nil
}

func (elem *Element) CachedAccessKey() (string, error) {
	var retVal *uint16

	hr, _, _ := syscall.SyscallN(
		elem.VTable().Get_CachedAccessKey,
		uintptr(unsafe.Pointer(elem)),
		uintptr(unsafe.Pointer(&retVal)),
	)

	if hr != 0 {
		return "", ole.NewError(hr)
	}

	return ole.BstrToString(retVal), nil
}

func (elem *Element) CachedHasKeyboardFocus() (bool, error) {
	var retVal bool

	hr, _, _ := syscall.SyscallN(
		elem.VTable().Get_CachedHasKeyboardFocus,
		uintptr(unsafe.Pointer(elem)),
		uintptr(unsafe.Pointer(&retVal)),
	)

	if hr != 0 {
		return false, ole.NewError(hr)
	}

	return retVal, nil
}

func (elem *Element) CachedIsKeyboardFocusable() (bool, error) {
	var retVal bool

	hr, _, _ := syscall.SyscallN(
		elem.VTable().Get_CachedIsKeyboardFocusable,
		uintptr(unsafe.Pointer(elem)),
		uintptr(unsafe.Pointer(&retVal)),
	)

	if hr != 0 {
		return false, ole.NewError(hr)
	}

	return retVal, nil
}

func (elem *Element) CachedIsEnabled() (bool, error) {
	var retVal bool

	hr, _, _ := syscall.SyscallN(
		elem.VTable().Get_CachedIsEnabled,
		uintptr(unsafe.Pointer(elem)),
		uintptr(unsafe.Pointer(&retVal)),
	)

	if hr != 0 {
		return false, ole.NewError(hr)
	}

	return retVal, nil
}

func (elem *Element) CachedAutomationId() (string, error) {
	var retVal *uint16

	hr, _, _ := syscall.SyscallN(
		elem.VTable().Get_CachedAutomationId,
		uintptr(unsafe.Pointer(elem)),
		uintptr(unsafe.Pointer(&retVal)),
	)

	if hr != 0 {
		return "", ole.NewError(hr)
	}

	return ole.BstrToString(retVal), nil
}

func (elem *Element) CachedClassName() (string, error) {
	var retVal *uint16

	hr, _, _ := syscall.SyscallN(
		elem.VTable().Get_CachedClassName,
		uintptr(unsafe.Pointer(elem)),
		uintptr(unsafe.Pointer(&retVal)),
	)

	if hr != 0 {
		return "", ole.NewError(hr)
	}

	return ole.BstrToString(retVal), nil
}

func (elem *Element) CachedHelpText() (string, error) {
	var retVal *uint16

	hr, _, _ := syscall.SyscallN(
		elem.VTable().Get_CachedHelpText,
		uintptr(unsafe.Pointer(elem)),
		uintptr(unsafe.Pointer(&retVal)),
	)

	if hr != 0 {
		return "", ole.NewError(hr)
	}

	return ole.BstrToString(retVal), nil
}

func (elem *Element) CachedCulture() (int32, error) {
	var retVal int32

	hr, _, _ := syscall.SyscallN(
		elem.VTable().Get_CachedCulture,
		uintptr(unsafe.Pointer(elem)),
		uintptr(unsafe.Pointer(&retVal)),
	)

	if hr != 0 {
		return 0, ole.NewError(hr)
	}

	return retVal, nil
}

func (elem *Element) CachedIsControlElement() (bool, error) {
	var retVal bool

	hr, _, _ := syscall.SyscallN(
		elem.VTable().Get_CachedIsControlElement,
		uintptr(unsafe.Pointer(elem)),
		uintptr(unsafe.Pointer(&retVal)),
	)

	if hr != 0 {
		return false, ole.NewError(hr)
	}

	return retVal, nil
}

func (elem *Element) CachedIsContentElement() (bool, error) {
	var retVal bool

	hr, _, _ := syscall.SyscallN(
		elem.VTable().Get_CachedIsContentElement,
		uintptr(unsafe.Pointer(elem)),
		uintptr(unsafe.Pointer(&retVal)),
	)

	if hr != 0 {
		return false, ole.NewError(hr)
	}

	return retVal, nil
}

func (elem *Element) CachedIsPassword() (bool, error) {
	var retVal bool

	hr, _, _ := syscall.SyscallN(
		elem.VTable().Get_CachedIsPassword,
		uintptr(unsafe.Pointer(elem)),
		uintptr(unsafe.Pointer(&retVal)),
	)

	if hr != 0 {
		return false, ole.NewError(hr)
	}

	return retVal, nil
}

func (elem *Element) CachedNativeWindowHandle() (syscall.Handle, error) {
	var retVal syscall.Handle

	hr, _, _ := syscall.SyscallN(
		elem.VTable().Get_CachedNativeWindowHandle,
		uintptr(unsafe.Pointer(elem)),
		uintptr(unsafe.Pointer(&retVal)),
	)

	if hr != 0 {
		return 0, ole.NewError(hr)
	}

	return retVal, nil
}

func (elem *Element) CachedItemType() (string, error) {
	var retVal *uint16

	hr, _, _ := syscall.SyscallN(
		elem.VTable().Get_CachedItemType,
		uintptr(unsafe.Pointer(elem)),
		uintptr(unsafe.Pointer(&retVal)),
	)

	if hr != 0 {
		return "", ole.NewError(hr)
	}

	return ole.BstrToString(retVal), nil
}

func (elem *Element) CachedIsOffscreen() (bool, error) {
	var retVal bool

	hr, _, _ := syscall.SyscallN(
		elem.VTable().Get_CachedIsOffscreen,
		uintptr(unsafe.Pointer(elem)),
		uintptr(unsafe.Pointer(&retVal)),
	)

	if hr != 0 {
		return false, ole.NewError(hr)
	}

	return retVal, nil
}

func (elem *Element) CachedOrientation() (OrientationType, error) {
	var retVal OrientationType

	hr, _, _ := syscall.SyscallN(
		elem.VTable().Get_CachedOrientation,
		uintptr(unsafe.Pointer(elem)),
		uintptr(unsafe.Pointer(&retVal)),
	)

	if hr != 0 {
		return 0, ole.NewError(hr)
	}

	return retVal, nil
}

func (elem *Element) CachedFrameworkId() (string, error) {
	var retVal *uint16

	hr, _, _ := syscall.SyscallN(
		elem.VTable().Get_CachedFrameworkId,
		uintptr(unsafe.Pointer(elem)),
		uintptr(unsafe.Pointer(&retVal)),
	)

	if hr != 0 {
		return "", ole.NewError(hr)
	}

	return ole.BstrToString(retVal), nil
}

func (elem *Element) CachedIsRequiredForForm() (bool, error) {
	var retVal bool

	hr, _, _ := syscall.SyscallN(
		elem.VTable().Get_CachedIsRequiredForForm,
		uintptr(unsafe.Pointer(elem)),
		uintptr(unsafe.Pointer(&retVal)),
	)

	if hr != 0 {
		return false, ole.NewError(hr)
	}

	return retVal, nil
}

func (elem *Element) CachedItemStatus() (string, error) {
	var retVal *uint16

	hr, _, _ := syscall.SyscallN(
		elem.VTable().Get_CachedItemStatus,
		uintptr(unsafe.Pointer(elem)),
		uintptr(unsafe.Pointer(&retVal)),
	)

	if hr != 0 {
		return "", ole.NewError(hr)
	}

	return ole.BstrToString(retVal), nil
}

func (elem *Element) CachedBoundingRectangle() (Rect, error) {
	var retVal Rect

	hr, _, _ := syscall.SyscallN(
		elem.VTable().Get_CachedBoundingRectangle,
		uintptr(unsafe.Pointer(elem)),
		uintptr(unsafe.Pointer(&retVal)),
	)

	if hr != 0 {
		return Rect{}, ole.NewError(hr)
	}

	return retVal, nil
}

func (elem *Element) CachedLabeledBy() (*Element, error) {
	var retVal *Element

	hr, _, _ := syscall.SyscallN(
		elem.VTable().Get_CachedLabeledBy,
		uintptr(unsafe.Pointer(elem)),
		uintptr(unsafe.Pointer(&retVal)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return retVal, nil
}

func (elem *Element) CachedAriaRole() (string, error) {
	var retVal *uint16

	hr, _, _ := syscall.SyscallN(
		elem.VTable().Get_CachedAriaRole,
		uintptr(unsafe.Pointer(elem)),
		uintptr(unsafe.Pointer(&retVal)),
	)

	if hr != 0 {
		return "", ole.NewError(hr)
	}

	return ole.BstrToString(retVal), nil
}

func (elem *Element) CachedAriaProperties() (string, error) {
	var retVal *uint16

	hr, _, _ := syscall.SyscallN(
		elem.VTable().Get_CachedAriaProperties,
		uintptr(unsafe.Pointer(elem)),
		uintptr(unsafe.Pointer(&retVal)),
	)

	if hr != 0 {
		return "", ole.NewError(hr)
	}

	return ole.BstrToString(retVal), nil
}

func (elem *Element) CachedIsDataValidForForm() (bool, error) {
	var retVal bool

	hr, _, _ := syscall.SyscallN(
		elem.VTable().Get_CachedIsDataValidForForm,
		uintptr(unsafe.Pointer(elem)),
		uintptr(unsafe.Pointer(&retVal)),
	)

	if hr != 0 {
		return false, ole.NewError(hr)
	}

	return retVal, nil
}

func (elem *Element) CachedControllerFor() (*ElementArray, error) {
	var retVal *ElementArray

	hr, _, _ := syscall.SyscallN(
		elem.VTable().Get_CachedControllerFor,
		uintptr(unsafe.Pointer(elem)),
		uintptr(unsafe.Pointer(&retVal)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return retVal, nil
}

func (elem *Element) CachedDescribedBy() (*ElementArray, error) {
	var retVal *ElementArray

	hr, _, _ := syscall.SyscallN(
		elem.VTable().Get_CachedDescribedBy,
		uintptr(unsafe.Pointer(elem)),
		uintptr(unsafe.Pointer(&retVal)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return retVal, nil
}

func (elem *Element) CachedFlowsTo() (*ElementArray, error) {
	var retVal *ElementArray

	hr, _, _ := syscall.SyscallN(
		elem.VTable().Get_CachedFlowsTo,
		uintptr(unsafe.Pointer(elem)),
		uintptr(unsafe.Pointer(&retVal)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return retVal, nil
}

func (elem *Element) CachedProviderDescription() (string, error) {
	var retVal *uint16

	hr, _, _ := syscall.SyscallN(
		elem.VTable().Get_CachedProviderDescription,
		uintptr(unsafe.Pointer(elem)),
		uintptr(unsafe.Pointer(&retVal)),
	)

	if hr != 0 {
		return "", ole.NewError(hr)
	}

	return ole.BstrToString(retVal), nil
}

func (elem *Element) ClickablePoint() (Point, bool, error) {
	var clickable Point
	var gotClickable bool

	hr, _, _ := syscall.SyscallN(
		elem.VTable().GetClickablePoint,
		uintptr(unsafe.Pointer(elem)),
		uintptr(unsafe.Pointer(&clickable)),
		uintptr(unsafe.Pointer(&gotClickable)),
	)

	if hr != 0 {
		return Point{}, false, ole.NewError(hr)
	}

	return clickable, gotClickable, nil
}

func (elem *Element) GetInvokePattern() (*InvokePattern, error) {
	if elem == nil {
		return nil, ErrElementIsNil
	}

	patternObject, err := elem.GetCurrentPattern(InvokePatternId)

	if err != nil {
		return nil, err
	}

	if patternObject == nil {
		return nil, ErrPatternNotFound
	}

	invokePattern, err := patternObject.QueryInterface(IID_IUIAutomationInvokePattern)

	if err != nil {
		return nil, err
	}

	return (*InvokePattern)(unsafe.Pointer(invokePattern)), nil
}

func (elem *Element) GetSelectionPattern() (*SelectionPattern, error) {
	if elem == nil {
		return nil, ErrElementIsNil
	}

	patternObject, err := elem.GetCurrentPattern(SelectionPatternId)

	if err != nil {
		return nil, err
	}

	if patternObject == nil {
		return nil, ErrPatternNotFound
	}

	selectionPattern, err := patternObject.QueryInterface(IID_IUIAutomationSelectionPattern)

	if err != nil {
		return nil, err
	}

	return (*SelectionPattern)(unsafe.Pointer(selectionPattern)), nil
}

func (elem *Element) GetValuePattern() (*ValuePattern, error) {
	if elem == nil {
		return nil, ErrElementIsNil
	}

	patternObject, err := elem.GetCurrentPattern(ValuePatternId)

	if err != nil {
		return nil, err
	}

	if patternObject == nil {
		return nil, ErrPatternNotFound
	}

	valuePattern, err := patternObject.QueryInterface(IID_IUIAutomationValuePattern)

	if err != nil {
		return nil, err
	}

	return (*ValuePattern)(unsafe.Pointer(valuePattern)), nil
}

func (elem *Element) GetRangeValuePattern() (*RangeValuePattern, error) {
	if elem == nil {
		return nil, ErrElementIsNil
	}

	patternObject, err := elem.GetCurrentPattern(RangeValuePatternId)

	if err != nil {
		return nil, err
	}

	if patternObject == nil {
		return nil, ErrPatternNotFound
	}

	rangeValuePattern, err := patternObject.QueryInterface(IID_IUIAutomationRangeValuePattern)

	if err != nil {
		return nil, err
	}

	return (*RangeValuePattern)(unsafe.Pointer(rangeValuePattern)), nil
}

func (elem *Element) GetScrollPattern() (*ScrollPattern, error) {
	if elem == nil {
		return nil, ErrElementIsNil
	}

	patternObject, err := elem.GetCurrentPattern(ScrollPatternId)

	if err != nil {
		return nil, err
	}

	if patternObject == nil {
		return nil, ErrPatternNotFound
	}

	scrollPattern, err := patternObject.QueryInterface(IID_IUIAutomationScrollPattern)

	if err != nil {
		return nil, err
	}

	return (*ScrollPattern)(unsafe.Pointer(scrollPattern)), nil
}

func (elem *Element) GetExpandCollapsePattern() (*ExpandCollapsePattern, error) {
	if elem == nil {
		return nil, ErrElementIsNil
	}

	patternObject, err := elem.GetCurrentPattern(ExpandCollapsePatternId)

	if err != nil {
		return nil, err
	}

	if patternObject == nil {
		return nil, ErrPatternNotFound
	}

	expandCollapsePattern, err := patternObject.QueryInterface(IID_IUIAutomationExpandCollapsePattern)

	if err != nil {
		return nil, err
	}

	return (*ExpandCollapsePattern)(unsafe.Pointer(expandCollapsePattern)), nil
}

func (elem *Element) GetGridPattern() (*GridPattern, error) {
	if elem == nil {
		return nil, ErrElementIsNil
	}

	patternObject, err := elem.GetCurrentPattern(GridPatternId)

	if err != nil {
		return nil, err
	}

	if patternObject == nil {
		return nil, ErrPatternNotFound
	}

	gridPattern, err := patternObject.QueryInterface(IID_IUIAutomationGridPattern)

	if err != nil {
		return nil, err
	}

	return (*GridPattern)(unsafe.Pointer(gridPattern)), nil
}

func (elem *Element) GetGridItemPattern() (*GridItemPattern, error) {
	if elem == nil {
		return nil, ErrElementIsNil
	}

	patternObject, err := elem.GetCurrentPattern(GridItemPatternId)

	if err != nil {
		return nil, err
	}

	if patternObject == nil {
		return nil, ErrPatternNotFound
	}

	gridItemPattern, err := patternObject.QueryInterface(IID_IUIAutomationGridItemPattern)

	if err != nil {
		return nil, err
	}

	return (*GridItemPattern)(unsafe.Pointer(gridItemPattern)), nil
}

func (elem *Element) GetMultipleViewPattern() (*MultipleViewPattern, error) {
	if elem == nil {
		return nil, ErrElementIsNil
	}

	patternObject, err := elem.GetCurrentPattern(MultipleViewPatternId)

	if err != nil {
		return nil, err
	}

	if patternObject == nil {
		return nil, ErrPatternNotFound
	}

	multipleViewPattern, err := patternObject.QueryInterface(IID_IUIAutomationMultipleViewPattern)

	if err != nil {
		return nil, err
	}

	return (*MultipleViewPattern)(unsafe.Pointer(multipleViewPattern)), nil
}

func (elem *Element) GetWindowPattern() (*WindowPattern, error) {
	if elem == nil {
		return nil, ErrElementIsNil
	}

	patternObject, err := elem.GetCurrentPattern(WindowPatternId)

	if err != nil {
		return nil, err
	}

	if patternObject == nil {
		return nil, ErrPatternNotFound
	}

	windowPattern, err := patternObject.QueryInterface(IID_IUIAutomationWindowPattern)

	if err != nil {
		return nil, err
	}

	return (*WindowPattern)(unsafe.Pointer(windowPattern)), nil
}

func (elem *Element) GetSelectionItemPattern() (*SelectionItemPattern, error) {
	if elem == nil {
		return nil, ErrElementIsNil
	}

	patternObject, err := elem.GetCurrentPattern(SelectionItemPatternId)

	if err != nil {
		return nil, err
	}

	if patternObject == nil {
		return nil, ErrPatternNotFound
	}

	selectionItemPattern, err := patternObject.QueryInterface(IID_IUIAutomationSelectionItemPattern)

	if err != nil {
		return nil, err
	}

	return (*SelectionItemPattern)(unsafe.Pointer(selectionItemPattern)), nil
}

func (elem *Element) GetDockPattern() (*DockPattern, error) {
	if elem == nil {
		return nil, ErrElementIsNil
	}

	patternObject, err := elem.GetCurrentPattern(DockPatternId)

	if err != nil {
		return nil, err
	}

	if patternObject == nil {
		return nil, ErrPatternNotFound
	}

	dockPattern, err := patternObject.QueryInterface(IID_IUIAutomationDockPattern)

	if err != nil {
		return nil, err
	}

	return (*DockPattern)(unsafe.Pointer(dockPattern)), nil
}

func (elem *Element) GetTablePattern() (*TablePattern, error) {
	if elem == nil {
		return nil, ErrElementIsNil
	}

	patternObject, err := elem.GetCurrentPattern(TablePatternId)

	if err != nil {
		return nil, err
	}

	if patternObject == nil {
		return nil, ErrPatternNotFound
	}

	tablePattern, err := patternObject.QueryInterface(IID_IUIAutomationTablePattern)

	if err != nil {
		return nil, err
	}

	return (*TablePattern)(unsafe.Pointer(tablePattern)), nil
}

func (elem *Element) GetTableItemPattern() (*TableItemPattern, error) {
	if elem == nil {
		return nil, ErrElementIsNil
	}

	patternObject, err := elem.GetCurrentPattern(TableItemPatternId)

	if err != nil {
		return nil, err
	}

	if patternObject == nil {
		return nil, ErrPatternNotFound
	}

	tableItemPattern, err := patternObject.QueryInterface(IID_IUIAutomationTableItemPattern)

	if err != nil {
		return nil, err
	}

	return (*TableItemPattern)(unsafe.Pointer(tableItemPattern)), nil
}

func (elem *Element) GetTextPattern() (*TextPattern, error) {
	if elem == nil {
		return nil, ErrElementIsNil
	}

	patternObject, err := elem.GetCurrentPattern(TextPatternId)

	if err != nil {
		return nil, err
	}

	if patternObject == nil {
		return nil, ErrPatternNotFound
	}

	textPattern, err := patternObject.QueryInterface(IID_IUIAutomationTextPattern)

	if err != nil {
		return nil, err
	}

	return (*TextPattern)(unsafe.Pointer(textPattern)), nil
}

func (elem *Element) GetTogglePattern() (*TogglePattern, error) {
	if elem == nil {
		return nil, ErrElementIsNil
	}

	patternObject, err := elem.GetCurrentPattern(TogglePatternId)

	if err != nil {
		return nil, err
	}

	if patternObject == nil {
		return nil, ErrPatternNotFound
	}

	togglePattern, err := patternObject.QueryInterface(IID_IUIAutomationTogglePattern)

	if err != nil {
		return nil, err
	}

	return (*TogglePattern)(unsafe.Pointer(togglePattern)), nil
}

func (elem *Element) GetTransformPattern() (*TransformPattern, error) {
	if elem == nil {
		return nil, ErrElementIsNil
	}

	patternObject, err := elem.GetCurrentPattern(TransformPatternId)

	if err != nil {
		return nil, err
	}

	if patternObject == nil {
		return nil, ErrPatternNotFound
	}

	transformPattern, err := patternObject.QueryInterface(IID_IUIAutomationTransformPattern)

	if err != nil {
		return nil, err
	}

	return (*TransformPattern)(unsafe.Pointer(transformPattern)), nil
}

func (elem *Element) GetScrollItemPattern() (*ScrollItemPattern, error) {
	if elem == nil {
		return nil, ErrElementIsNil
	}

	patternObject, err := elem.GetCurrentPattern(ScrollItemPatternId)

	if err != nil {
		return nil, err
	}

	if patternObject == nil {
		return nil, ErrPatternNotFound
	}

	scrollItemPattern, err := patternObject.QueryInterface(IID_IUIAutomationScrollItemPattern)

	if err != nil {
		return nil, err
	}

	return (*ScrollItemPattern)(unsafe.Pointer(scrollItemPattern)), nil
}

func (elem *Element) GetLegacyAccessiblePattern() (*LegacyAccessiblePattern, error) {
	if elem == nil {
		return nil, ErrElementIsNil
	}

	patternObject, err := elem.GetCurrentPattern(LegacyIAccessiblePatternId)

	if err != nil {
		return nil, err
	}

	if patternObject == nil {
		return nil, ErrPatternNotFound
	}

	legacyAccessiblePattern, err := patternObject.QueryInterface(IID_IUIAutomationLegacyIAccessiblePattern)

	if err != nil {
		return nil, err
	}

	return (*LegacyAccessiblePattern)(unsafe.Pointer(legacyAccessiblePattern)), nil
}

func (elem *Element) GetItemContainerPattern() (*ItemContainerPattern, error) {
	if elem == nil {
		return nil, ErrElementIsNil
	}

	patternObject, err := elem.GetCurrentPattern(ItemContainerPatternId)

	if err != nil {
		return nil, err
	}

	if patternObject == nil {
		return nil, ErrPatternNotFound
	}

	itemContainerPattern, err := patternObject.QueryInterface(IID_IUIAutomationItemContainerPattern)

	if err != nil {
		return nil, err
	}

	return (*ItemContainerPattern)(unsafe.Pointer(itemContainerPattern)), nil
}

func (elem *Element) GetVirtualizedItemPattern() (*VirtualizedItemPattern, error) {
	if elem == nil {
		return nil, ErrElementIsNil
	}

	patternObject, err := elem.GetCurrentPattern(VirtualizedItemPatternId)

	if err != nil {
		return nil, err
	}

	if patternObject == nil {
		return nil, ErrPatternNotFound
	}

	virtualizedItemPattern, err := patternObject.QueryInterface(IID_IUIAutomationVirtualizedItemPattern)

	if err != nil {
		return nil, err
	}

	return (*VirtualizedItemPattern)(unsafe.Pointer(virtualizedItemPattern)), nil
}

func (elem *Element) GetSynchronizedInputPattern() (*SynchronizedInputPattern, error) {
	if elem == nil {
		return nil, ErrElementIsNil
	}

	patternObject, err := elem.GetCurrentPattern(SynchronizedInputPatternId)

	if err != nil {
		return nil, err
	}

	if patternObject == nil {
		return nil, ErrPatternNotFound
	}

	synchronizedInputPattern, err := patternObject.QueryInterface(IID_IUIAutomationSynchronizedInputPattern)

	if err != nil {
		return nil, err
	}

	return (*SynchronizedInputPattern)(unsafe.Pointer(synchronizedInputPattern)), nil
}

func (elem *Element) GetObjectModelPattern() (*ObjectModelPattern, error) {
	if elem == nil {
		return nil, ErrElementIsNil
	}

	patternObject, err := elem.GetCurrentPattern(ObjectModelPatternId)

	if err != nil {
		return nil, err
	}

	if patternObject == nil {
		return nil, ErrPatternNotFound
	}

	objectModelPattern, err := patternObject.QueryInterface(IID_IUIAutomationObjectModelPattern)

	if err != nil {
		return nil, err
	}

	return (*ObjectModelPattern)(unsafe.Pointer(objectModelPattern)), nil
}

func (elem *Element) GetAnnotationPattern() (*AnnotationPattern, error) {
	if elem == nil {
		return nil, ErrElementIsNil
	}

	patternObject, err := elem.GetCurrentPattern(AnnotationPatternId)

	if err != nil {
		return nil, err
	}

	if patternObject == nil {
		return nil, ErrPatternNotFound
	}

	annotationPattern, err := patternObject.QueryInterface(IID_IUIAutomationAnnotationPattern)

	if err != nil {
		return nil, err
	}

	return (*AnnotationPattern)(unsafe.Pointer(annotationPattern)), nil
}

func (elem *Element) GetTextPattern2() (*TextPattern2, error) {
	if elem == nil {
		return nil, ErrElementIsNil
	}

	patternObject, err := elem.GetCurrentPattern(TextPattern2Id)

	if err != nil {
		return nil, err
	}

	if patternObject == nil {
		return nil, ErrPatternNotFound
	}

	textPattern2, err := patternObject.QueryInterface(IID_IUIAutomationTextPattern2)

	if err != nil {
		return nil, err
	}

	return (*TextPattern2)(unsafe.Pointer(textPattern2)), nil
}

func (elem *Element) GetStylesPattern() (*StylesPattern, error) {
	if elem == nil {
		return nil, ErrElementIsNil
	}

	patternObject, err := elem.GetCurrentPattern(StylesPatternId)

	if err != nil {
		return nil, err
	}

	if patternObject == nil {
		return nil, ErrPatternNotFound
	}

	stylesPattern, err := patternObject.QueryInterface(IID_IUIAutomationStylesPattern)

	if err != nil {
		return nil, err
	}

	return (*StylesPattern)(unsafe.Pointer(stylesPattern)), nil
}

func (elem *Element) GetSpreadsheetPattern() (*SpreadsheetPattern, error) {
	if elem == nil {
		return nil, ErrElementIsNil
	}

	patternObject, err := elem.GetCurrentPattern(SpreadsheetPatternId)

	if err != nil {
		return nil, err
	}

	if patternObject == nil {
		return nil, ErrPatternNotFound
	}

	spreadsheetPattern, err := patternObject.QueryInterface(IID_IUIAutomationSpreadsheetPattern)

	if err != nil {
		return nil, err
	}

	return (*SpreadsheetPattern)(unsafe.Pointer(spreadsheetPattern)), nil
}

func (elem *Element) GetSpreadsheetItemPattern() (*SpreadsheetItemPattern, error) {
	if elem == nil {
		return nil, ErrElementIsNil
	}

	patternObject, err := elem.GetCurrentPattern(SpreadsheetItemPatternId)

	if err != nil {
		return nil, err
	}

	if patternObject == nil {
		return nil, ErrPatternNotFound
	}

	spreadsheetItemPattern, err := patternObject.QueryInterface(IID_IUIAutomationSpreadsheetItemPattern)

	if err != nil {
		return nil, err
	}

	return (*SpreadsheetItemPattern)(unsafe.Pointer(spreadsheetItemPattern)), nil
}

func (elem *Element) GetTransformPattern2() (*TransformPattern2, error) {
	if elem == nil {
		return nil, ErrElementIsNil
	}

	patternObject, err := elem.GetCurrentPattern(TransformPattern2Id)

	if err != nil {
		return nil, err
	}

	if patternObject == nil {
		return nil, ErrPatternNotFound
	}

	transformPattern2, err := patternObject.QueryInterface(IID_IUIAutomationTransformPattern2)

	if err != nil {
		return nil, err
	}

	return (*TransformPattern2)(unsafe.Pointer(transformPattern2)), nil
}

func (elem *Element) GetTextChildPattern() (*TextChildPattern, error) {
	if elem == nil {
		return nil, ErrElementIsNil
	}

	patternObject, err := elem.GetCurrentPattern(TextChildPatternId)

	if err != nil {
		return nil, err
	}

	if patternObject == nil {
		return nil, ErrPatternNotFound
	}

	textChildPattern, err := patternObject.QueryInterface(IID_IUIAutomationTextChildPattern)

	if err != nil {
		return nil, err
	}

	return (*TextChildPattern)(unsafe.Pointer(textChildPattern)), nil
}

func (elem *Element) GetDragPattern() (*DragPattern, error) {
	if elem == nil {
		return nil, ErrElementIsNil
	}

	patternObject, err := elem.GetCurrentPattern(DragPatternId)

	if err != nil {
		return nil, err
	}

	if patternObject == nil {
		return nil, ErrPatternNotFound
	}

	dragPattern, err := patternObject.QueryInterface(IID_IUIAutomationDragPattern)

	if err != nil {
		return nil, err
	}

	return (*DragPattern)(unsafe.Pointer(dragPattern)), nil
}

func (elem *Element) GetDropTargetPattern() (*DropTargetPattern, error) {
	if elem == nil {
		return nil, ErrElementIsNil
	}

	patternObject, err := elem.GetCurrentPattern(DropTargetPatternId)

	if err != nil {
		return nil, err
	}

	if patternObject == nil {
		return nil, ErrPatternNotFound
	}

	dropTargetPattern, err := patternObject.QueryInterface(IID_IUIAutomationDropTargetPattern)

	if err != nil {
		return nil, err
	}

	return (*DropTargetPattern)(unsafe.Pointer(dropTargetPattern)), nil
}

func (elem *Element) GetTextEditPattern() (*TextEditPattern, error) {
	if elem == nil {
		return nil, ErrElementIsNil
	}

	patternObject, err := elem.GetCurrentPattern(TextEditPatternId)

	if err != nil {
		return nil, err
	}

	if patternObject == nil {
		return nil, ErrPatternNotFound
	}

	textEditPattern, err := patternObject.QueryInterface(IID_IUIAutomationTextEditPattern)

	if err != nil {
		return nil, err
	}

	return (*TextEditPattern)(unsafe.Pointer(textEditPattern)), nil
}

func (elem *Element) GetCustomNavigationPattern() (*CustomNavigationPattern, error) {
	if elem == nil {
		return nil, ErrElementIsNil
	}

	patternObject, err := elem.GetCurrentPattern(CustomNavigationPatternId)

	if err != nil {
		return nil, err
	}

	if patternObject == nil {
		return nil, ErrPatternNotFound
	}

	customNavigationPattern, err := patternObject.QueryInterface(IID_IUIAutomationCustomNavigationPattern)

	if err != nil {
		return nil, err
	}

	return (*CustomNavigationPattern)(unsafe.Pointer(customNavigationPattern)), nil
}

func (elem *Element) SetValue(value string) error {
	if elem == nil {
		return ErrElementIsNil
	}

	valuePattern, err := elem.GetValuePattern()

	if err != nil {
		return err
	}

	defer valuePattern.Release()

	return valuePattern.SetValue(value)
}

func (elem *Element) Invoke() error {
	if elem == nil {
		return ErrElementIsNil
	}

	invokePattern, err := elem.GetInvokePattern()

	if err != nil {
		return err
	}

	defer invokePattern.Release()

	return invokePattern.Invoke()
}

func (elem *Element) DoDefaultAction() error {
	if elem == nil {
		return ErrElementIsNil
	}

	legacyAccessiblePattern, err := elem.GetLegacyAccessiblePattern()

	if err != nil {
		return err
	}

	defer legacyAccessiblePattern.Release()

	return legacyAccessiblePattern.DoDefaultAction()
}

func (elem *Element) NormalizeWindow() error {
	if elem == nil {
		return ErrElementIsNil
	}

	windowPattern, err := elem.GetWindowPattern()
	
	if err != nil {
		return err
	}
	
	defer windowPattern.Release()

	state, err := windowPattern.CurrentWindowVisualState()
	
	if err != nil {
		return fmt.Errorf("failed to get window visual state: %w", err)
	}

	if state == WindowVisualStateNormal {
		return nil // already normalized
	}

	return windowPattern.SetWindowVisualState(WindowVisualStateNormal)
}

func (elem *Element) CloseWindow() error {
	if elem == nil {
		return ErrElementIsNil
	}

	windowPattern, err := elem.GetWindowPattern()

	if err != nil {
		return err
	}

	defer windowPattern.Release()

	return windowPattern.Close()
}

func (elem *Element) CurrentControlTypeName() (string, error) {
	if elem == nil {
		return "", ErrElementIsNil
	}

	controlTypeId, err := elem.CurrentControlType()

	if err != nil {
		return "", err
	}

	controlTypeName := ControlTypeNameFromId(controlTypeId)

	return controlTypeName, nil
}

type ElementArray struct {
	ole.IUnknown
}

type ElementArrayVtbl struct {
	ole.IUnknownVtbl
	Get_Length uintptr
	GetElement uintptr
}

func (arr *ElementArray) VTable() *ElementArrayVtbl {
	return (*ElementArrayVtbl)(unsafe.Pointer(arr.RawVTable))
}

func (arr *ElementArray) Length() (int32, error) {
	var retVal int32

	hr, _, _ := syscall.SyscallN(
		arr.VTable().Get_Length,
		uintptr(unsafe.Pointer(arr)),
		uintptr(unsafe.Pointer(&retVal)),
	)

	if hr != 0 {
		return 0, ole.NewError(hr)
	}

	return retVal, nil
}

func (arr *ElementArray) GetElement(index int32) (*Element, error) {
	var retVal *Element

	hr, _, _ := syscall.SyscallN(
		arr.VTable().GetElement,
		uintptr(unsafe.Pointer(arr)),
		uintptr(index),
		uintptr(unsafe.Pointer(&retVal)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return retVal, nil
}
