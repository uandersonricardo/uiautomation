package uiautomation

import (
	"syscall"
	"unsafe"

	"github.com/go-ole/go-ole"
)

type UIAutomationElement struct {
	ole.IUnknown
}

type UIAutomationElementVtbl struct {
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

func (elem *UIAutomationElement) VTable() *UIAutomationElementVtbl {
	return (*UIAutomationElementVtbl)(unsafe.Pointer(elem.RawVTable))
}

func (elem *UIAutomationElement) SetFocus() error {
	hr, _, _ := syscall.SyscallN(
		elem.VTable().SetFocus,
		uintptr(unsafe.Pointer(elem)),
	)

	if hr != 0 {
		return ole.NewError(hr)
	}

	return nil
}

func (elem *UIAutomationElement) SetValue(value string) error {
	valuePattern, err := elem.GetValuePattern()
	if err != nil {
		return err
	}
	defer valuePattern.Release()
	return valuePattern.SetValue(value)
}

func (elem *UIAutomationElement) Invoke() error {
	invokePattern, err := elem.GetInvokePattern()
	if err != nil {
		return err
	}
	invokePattern.Invoke()
	invokePattern.Release()
	return nil
}

func (elem *UIAutomationElement) DoDefaultAction() error {
	legacyAccessiblePattern, err := elem.GetLegacyAccessiblePattern()
	if err != nil {
		return err
	}
	legacyAccessiblePattern.DoDefaultAction()
	legacyAccessiblePattern.Release()
	return nil
}

func (elem *UIAutomationElement) GetCurrentPattern(patternId PatternId) (*ole.IUnknown, error) {
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

func (elem *UIAutomationElement) GetInvokePattern() (*UIAutomationInvokePattern, error) {
	patternObject, err := elem.GetCurrentPattern(InvokePatternId)

	if err != nil {
		return nil, err
	}

	invokePattern, err := patternObject.QueryInterface(IID_IUIAutomationInvokePattern)

	if err != nil {
		return nil, err
	}

	return (*UIAutomationInvokePattern)(unsafe.Pointer(invokePattern)), nil
}

func (elem *UIAutomationElement) GetLegacyAccessiblePattern() (*UIAutomationLegacyAccessiblePattern, error) {
	patternObject, err := elem.GetCurrentPattern(LegacyIAccessiblePatternId)

	if err != nil {
		return nil, err
	}

	legacyAccessiblePattern, err := patternObject.QueryInterface(IID_IUIAutomationLegacyIAccessiblePattern)

	if err != nil {
		return nil, err
	}

	return (*UIAutomationLegacyAccessiblePattern)(unsafe.Pointer(legacyAccessiblePattern)), nil
}

func (elem *UIAutomationElement) GetTextPattern() (*UIAutomationTextPattern, error) {
	patternObject, err := elem.GetCurrentPattern(TextPatternId)

	if err != nil {
		return nil, err
	}

	textPattern, err := patternObject.QueryInterface(IID_IUIAutomationTextPattern)

	if err != nil {
		return nil, err
	}

	return (*UIAutomationTextPattern)(unsafe.Pointer(textPattern)), nil
}

func (elem *UIAutomationElement) GetValuePattern() (*UIAutomationValuePattern, error) {
	patternObject, err := elem.GetCurrentPattern(ValuePatternId)

	if err != nil {
		return nil, err
	}

	valuePattern, err := patternObject.QueryInterface(IID_IUIAutomationValuePattern)

	if err != nil {
		return nil, err
	}

	return (*UIAutomationValuePattern)(unsafe.Pointer(valuePattern)), nil
}

func (elem *UIAutomationElement) CurrentControlType() (ControlTypeId, error) {
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

func (elem *UIAutomationElement) CurrentControlTypeName() (string, error) {
	controlTypeId, err := elem.CurrentControlType()

	if err != nil {
		return "", err
	}

	controlTypeName := ControlTypeNameFromId(controlTypeId)

	return controlTypeName, nil
}

func (elem *UIAutomationElement) CurrentName() (string, error) {
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

func (elem *UIAutomationElement) CurrentAutomationId() (string, error) {
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

func (elem *UIAutomationElement) CurrentClassName() (string, error) {
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

func (elem *UIAutomationElement) CurrentNativeWindowHandle() (syscall.Handle, error) {
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

func (elem *UIAutomationElement) CurrentBoundingRectangle() (Rect, error) {
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

func (elem *UIAutomationElement) CurrentPropertyValue(propertyId PropertyId) (ole.VARIANT, error) {
	var retVal ole.VARIANT

	ole.VariantInit(&retVal)

	hr, _, _ := syscall.SyscallN(
		elem.VTable().GetCurrentPropertyValue,
		uintptr(unsafe.Pointer(elem)),
		uintptr(propertyId),
		uintptr(unsafe.Pointer(&retVal)),
	)

	if hr != 0 {
		return retVal, ole.NewError(hr)
	}

	return retVal, nil
}

func (elem *UIAutomationElement) FindFirst(scope TreeScope, condition *UIAutomationCondition) (*UIAutomationElement, error) {
	var found *UIAutomationElement

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

type UIAutomationElementArray struct {
	ole.IUnknown
}

type UIAutomationElementArrayVtbl struct {
	ole.IUnknownVtbl
	Get_Length uintptr
	GetElement uintptr
}

func (arr *UIAutomationElementArray) VTable() *UIAutomationElementArrayVtbl {
	return (*UIAutomationElementArrayVtbl)(unsafe.Pointer(arr.RawVTable))
}

func (arr *UIAutomationElementArray) Length() (int32, error) {
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

func (arr *UIAutomationElementArray) GetElement(index int32) (*UIAutomationElement, error) {
	var retVal *UIAutomationElement

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
