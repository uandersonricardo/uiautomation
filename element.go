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

func (elem *Element) SetValue(value string) error {
	valuePattern, err := elem.GetValuePattern()
	if err != nil {
		return err
	}
	defer valuePattern.Release()
	return valuePattern.SetValue(value)
}

func (elem *Element) Invoke() error {
	invokePattern, err := elem.GetInvokePattern()
	if err != nil {
		return err
	}
	defer invokePattern.Release()
	return invokePattern.Invoke()
}

func (elem *Element) DoDefaultAction() error {
	legacyAccessiblePattern, err := elem.GetLegacyAccessiblePattern()
	if err != nil {
		return err
	}
	defer legacyAccessiblePattern.Release()
	return legacyAccessiblePattern.DoDefaultAction()
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

func (elem *Element) CurrentPropertyValue(propertyId PropertyId) (ole.VARIANT, error) {
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

func (elem *Element) GetInvokePattern() (*InvokePattern, error) {
	patternObject, err := elem.GetCurrentPattern(InvokePatternId)

	if err != nil {
		return nil, err
	}

	invokePattern, err := patternObject.QueryInterface(IID_IUIAutomationInvokePattern)

	if err != nil {
		return nil, err
	}

	return (*InvokePattern)(unsafe.Pointer(invokePattern)), nil
}

func (elem *Element) GetSelectionPattern() (*SelectionPattern, error) {
	patternObject, err := elem.GetCurrentPattern(SelectionPatternId)

	if err != nil {
		return nil, err
	}

	selectionPattern, err := patternObject.QueryInterface(IID_IUIAutomationSelectionPattern)

	if err != nil {
		return nil, err
	}

	return (*SelectionPattern)(unsafe.Pointer(selectionPattern)), nil
}

func (elem *Element) GetValuePattern() (*ValuePattern, error) {
	patternObject, err := elem.GetCurrentPattern(ValuePatternId)

	if err != nil {
		return nil, err
	}

	valuePattern, err := patternObject.QueryInterface(IID_IUIAutomationValuePattern)

	if err != nil {
		return nil, err
	}

	return (*ValuePattern)(unsafe.Pointer(valuePattern)), nil
}

func (elem *Element) GetRangeValuePattern() (*RangeValuePattern, error) {
	patternObject, err := elem.GetCurrentPattern(RangeValuePatternId)

	if err != nil {
		return nil, err
	}

	rangeValuePattern, err := patternObject.QueryInterface(IID_IUIAutomationRangeValuePattern)

	if err != nil {
		return nil, err
	}

	return (*RangeValuePattern)(unsafe.Pointer(rangeValuePattern)), nil
}

func (elem *Element) GetScrollPattern() (*ScrollPattern, error) {
	patternObject, err := elem.GetCurrentPattern(ScrollPatternId)

	if err != nil {
		return nil, err
	}

	scrollPattern, err := patternObject.QueryInterface(IID_IUIAutomationScrollPattern)

	if err != nil {
		return nil, err
	}

	return (*ScrollPattern)(unsafe.Pointer(scrollPattern)), nil
}

func (elem *Element) GetExpandCollapsePattern() (*ExpandCollapsePattern, error) {
	patternObject, err := elem.GetCurrentPattern(ExpandCollapsePatternId)

	if err != nil {
		return nil, err
	}

	expandCollapsePattern, err := patternObject.QueryInterface(IID_IUIAutomationExpandCollapsePattern)

	if err != nil {
		return nil, err
	}

	return (*ExpandCollapsePattern)(unsafe.Pointer(expandCollapsePattern)), nil
}

func (elem *Element) GetWindowPattern() (*WindowPattern, error) {
	patternObject, err := elem.GetCurrentPattern(WindowPatternId)

	if err != nil {
		return nil, err
	}

	windowPattern, err := patternObject.QueryInterface(IID_IUIAutomationWindowPattern)

	if err != nil {
		return nil, err
	}

	return (*WindowPattern)(unsafe.Pointer(windowPattern)), nil
}

func (elem *Element) GetTextPattern() (*TextPattern, error) {
	patternObject, err := elem.GetCurrentPattern(TextPatternId)

	if err != nil {
		return nil, err
	}

	textPattern, err := patternObject.QueryInterface(IID_IUIAutomationTextPattern)

	if err != nil {
		return nil, err
	}

	return (*TextPattern)(unsafe.Pointer(textPattern)), nil
}

func (elem *Element) GetScrollItemPattern() (*ScrollItemPattern, error) {
	patternObject, err := elem.GetCurrentPattern(ScrollItemPatternId)

	if err != nil {
		return nil, err
	}

	scrollItemPattern, err := patternObject.QueryInterface(IID_IUIAutomationScrollItemPattern)

	if err != nil {
		return nil, err
	}

	return (*ScrollItemPattern)(unsafe.Pointer(scrollItemPattern)), nil
}

func (elem *Element) GetLegacyAccessiblePattern() (*LegacyAccessiblePattern, error) {
	patternObject, err := elem.GetCurrentPattern(LegacyIAccessiblePatternId)

	if err != nil {
		return nil, err
	}

	legacyAccessiblePattern, err := patternObject.QueryInterface(IID_IUIAutomationLegacyIAccessiblePattern)

	if err != nil {
		return nil, err
	}

	return (*LegacyAccessiblePattern)(unsafe.Pointer(legacyAccessiblePattern)), nil
}

func (elem *Element) CurrentControlTypeName() (string, error) {
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
