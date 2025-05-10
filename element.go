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

func (elem *Element) CurrentPropertyValue(propertyId PropertyId) (*ole.VARIANT, error) {
	var retVal *ole.VARIANT

	ole.VariantInit(retVal)

	hr, _, _ := syscall.SyscallN(
		elem.VTable().GetCurrentPropertyValue,
		uintptr(unsafe.Pointer(elem)),
		uintptr(propertyId),
		uintptr(unsafe.Pointer(retVal)),
	)

	if hr != 0 {
		return retVal, ole.NewError(hr)
	}

	return retVal, nil
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

func (elem *Element) GetGridPattern() (*GridPattern, error) {
	patternObject, err := elem.GetCurrentPattern(GridPatternId)

	if err != nil {
		return nil, err
	}

	gridPattern, err := patternObject.QueryInterface(IID_IUIAutomationGridPattern)

	if err != nil {
		return nil, err
	}

	return (*GridPattern)(unsafe.Pointer(gridPattern)), nil
}

func (elem *Element) GetGridItemPattern() (*GridItemPattern, error) {
	patternObject, err := elem.GetCurrentPattern(GridItemPatternId)

	if err != nil {
		return nil, err
	}

	gridItemPattern, err := patternObject.QueryInterface(IID_IUIAutomationGridItemPattern)

	if err != nil {
		return nil, err
	}

	return (*GridItemPattern)(unsafe.Pointer(gridItemPattern)), nil
}

func (elem *Element) GetMultipleViewPattern() (*MultipleViewPattern, error) {
	patternObject, err := elem.GetCurrentPattern(MultipleViewPatternId)

	if err != nil {
		return nil, err
	}

	multipleViewPattern, err := patternObject.QueryInterface(IID_IUIAutomationMultipleViewPattern)

	if err != nil {
		return nil, err
	}

	return (*MultipleViewPattern)(unsafe.Pointer(multipleViewPattern)), nil
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

func (elem *Element) GetSelectionItemPattern() (*SelectionItemPattern, error) {
	patternObject, err := elem.GetCurrentPattern(SelectionItemPatternId)

	if err != nil {
		return nil, err
	}

	selectionItemPattern, err := patternObject.QueryInterface(IID_IUIAutomationSelectionItemPattern)

	if err != nil {
		return nil, err
	}

	return (*SelectionItemPattern)(unsafe.Pointer(selectionItemPattern)), nil
}

func (elem *Element) GetDockPattern() (*DockPattern, error) {
	patternObject, err := elem.GetCurrentPattern(DockPatternId)

	if err != nil {
		return nil, err
	}

	dockPattern, err := patternObject.QueryInterface(IID_IUIAutomationDockPattern)

	if err != nil {
		return nil, err
	}

	return (*DockPattern)(unsafe.Pointer(dockPattern)), nil
}

func (elem *Element) GetTablePattern() (*TablePattern, error) {
	patternObject, err := elem.GetCurrentPattern(TablePatternId)

	if err != nil {
		return nil, err
	}

	tablePattern, err := patternObject.QueryInterface(IID_IUIAutomationTablePattern)

	if err != nil {
		return nil, err
	}

	return (*TablePattern)(unsafe.Pointer(tablePattern)), nil
}

func (elem *Element) GetTableItemPattern() (*TableItemPattern, error) {
	patternObject, err := elem.GetCurrentPattern(TableItemPatternId)

	if err != nil {
		return nil, err
	}

	tableItemPattern, err := patternObject.QueryInterface(IID_IUIAutomationTableItemPattern)

	if err != nil {
		return nil, err
	}

	return (*TableItemPattern)(unsafe.Pointer(tableItemPattern)), nil
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

func (elem *Element) GetTogglePattern() (*TogglePattern, error) {
	patternObject, err := elem.GetCurrentPattern(TogglePatternId)

	if err != nil {
		return nil, err
	}

	togglePattern, err := patternObject.QueryInterface(IID_IUIAutomationTogglePattern)

	if err != nil {
		return nil, err
	}

	return (*TogglePattern)(unsafe.Pointer(togglePattern)), nil
}

func (elem *Element) GetTransformPattern() (*TransformPattern, error) {
	patternObject, err := elem.GetCurrentPattern(TransformPatternId)

	if err != nil {
		return nil, err
	}

	transformPattern, err := patternObject.QueryInterface(IID_IUIAutomationTransformPattern)

	if err != nil {
		return nil, err
	}

	return (*TransformPattern)(unsafe.Pointer(transformPattern)), nil
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

func (elem *Element) GetItemContainerPattern() (*ItemContainerPattern, error) {
	patternObject, err := elem.GetCurrentPattern(ItemContainerPatternId)

	if err != nil {
		return nil, err
	}

	itemContainerPattern, err := patternObject.QueryInterface(IID_IUIAutomationItemContainerPattern)

	if err != nil {
		return nil, err
	}

	return (*ItemContainerPattern)(unsafe.Pointer(itemContainerPattern)), nil
}

func (elem *Element) GetVirtualizedItemPattern() (*VirtualizedItemPattern, error) {
	patternObject, err := elem.GetCurrentPattern(VirtualizedItemPatternId)

	if err != nil {
		return nil, err
	}

	virtualizedItemPattern, err := patternObject.QueryInterface(IID_IUIAutomationVirtualizedItemPattern)

	if err != nil {
		return nil, err
	}

	return (*VirtualizedItemPattern)(unsafe.Pointer(virtualizedItemPattern)), nil
}

func (elem *Element) GetSynchronizedInputPattern() (*SynchronizedInputPattern, error) {
	patternObject, err := elem.GetCurrentPattern(SynchronizedInputPatternId)

	if err != nil {
		return nil, err
	}

	synchronizedInputPattern, err := patternObject.QueryInterface(IID_IUIAutomationSynchronizedInputPattern)

	if err != nil {
		return nil, err
	}

	return (*SynchronizedInputPattern)(unsafe.Pointer(synchronizedInputPattern)), nil
}

func (elem *Element) GetObjectModelPattern() (*ObjectModelPattern, error) {
	patternObject, err := elem.GetCurrentPattern(ObjectModelPatternId)

	if err != nil {
		return nil, err
	}

	objectModelPattern, err := patternObject.QueryInterface(IID_IUIAutomationObjectModelPattern)

	if err != nil {
		return nil, err
	}

	return (*ObjectModelPattern)(unsafe.Pointer(objectModelPattern)), nil
}

func (elem *Element) GetAnnotationPattern() (*AnnotationPattern, error) {
	patternObject, err := elem.GetCurrentPattern(AnnotationPatternId)

	if err != nil {
		return nil, err
	}

	annotationPattern, err := patternObject.QueryInterface(IID_IUIAutomationAnnotationPattern)

	if err != nil {
		return nil, err
	}

	return (*AnnotationPattern)(unsafe.Pointer(annotationPattern)), nil
}

func (elem *Element) GetTextPattern2() (*TextPattern2, error) {
	patternObject, err := elem.GetCurrentPattern(TextPattern2Id)

	if err != nil {
		return nil, err
	}

	textPattern2, err := patternObject.QueryInterface(IID_IUIAutomationTextPattern2)

	if err != nil {
		return nil, err
	}

	return (*TextPattern2)(unsafe.Pointer(textPattern2)), nil
}

func (elem *Element) GetStylesPattern() (*StylesPattern, error) {
	patternObject, err := elem.GetCurrentPattern(StylesPatternId)

	if err != nil {
		return nil, err
	}

	stylesPattern, err := patternObject.QueryInterface(IID_IUIAutomationStylesPattern)

	if err != nil {
		return nil, err
	}

	return (*StylesPattern)(unsafe.Pointer(stylesPattern)), nil
}

func (elem *Element) GetSpreadsheetPattern() (*SpreadsheetPattern, error) {
	patternObject, err := elem.GetCurrentPattern(SpreadsheetPatternId)

	if err != nil {
		return nil, err
	}

	spreadsheetPattern, err := patternObject.QueryInterface(IID_IUIAutomationSpreadsheetPattern)

	if err != nil {
		return nil, err
	}

	return (*SpreadsheetPattern)(unsafe.Pointer(spreadsheetPattern)), nil
}

func (elem *Element) GetSpreadsheetItemPattern() (*SpreadsheetItemPattern, error) {
	patternObject, err := elem.GetCurrentPattern(SpreadsheetItemPatternId)

	if err != nil {
		return nil, err
	}

	spreadsheetItemPattern, err := patternObject.QueryInterface(IID_IUIAutomationSpreadsheetItemPattern)

	if err != nil {
		return nil, err
	}

	return (*SpreadsheetItemPattern)(unsafe.Pointer(spreadsheetItemPattern)), nil
}

func (elem *Element) GetTransformPattern2() (*TransformPattern2, error) {
	patternObject, err := elem.GetCurrentPattern(TransformPattern2Id)

	if err != nil {
		return nil, err
	}

	transformPattern2, err := patternObject.QueryInterface(IID_IUIAutomationTransformPattern2)

	if err != nil {
		return nil, err
	}

	return (*TransformPattern2)(unsafe.Pointer(transformPattern2)), nil
}

func (elem *Element) GetTextChildPattern() (*TextChildPattern, error) {
	patternObject, err := elem.GetCurrentPattern(TextChildPatternId)

	if err != nil {
		return nil, err
	}

	textChildPattern, err := patternObject.QueryInterface(IID_IUIAutomationTextChildPattern)

	if err != nil {
		return nil, err
	}

	return (*TextChildPattern)(unsafe.Pointer(textChildPattern)), nil
}

func (elem *Element) GetDragPattern() (*DragPattern, error) {
	patternObject, err := elem.GetCurrentPattern(DragPatternId)

	if err != nil {
		return nil, err
	}

	dragPattern, err := patternObject.QueryInterface(IID_IUIAutomationDragPattern)

	if err != nil {
		return nil, err
	}

	return (*DragPattern)(unsafe.Pointer(dragPattern)), nil
}

func (elem *Element) GetDropTargetPattern() (*DropTargetPattern, error) {
	patternObject, err := elem.GetCurrentPattern(DropTargetPatternId)

	if err != nil {
		return nil, err
	}

	dropTargetPattern, err := patternObject.QueryInterface(IID_IUIAutomationDropTargetPattern)

	if err != nil {
		return nil, err
	}

	return (*DropTargetPattern)(unsafe.Pointer(dropTargetPattern)), nil
}

func (elem *Element) GetTextEditPattern() (*TextEditPattern, error) {
	patternObject, err := elem.GetCurrentPattern(TextEditPatternId)

	if err != nil {
		return nil, err
	}

	textEditPattern, err := patternObject.QueryInterface(IID_IUIAutomationTextEditPattern)

	if err != nil {
		return nil, err
	}

	return (*TextEditPattern)(unsafe.Pointer(textEditPattern)), nil
}

func (elem *Element) GetCustomNavigationPattern() (*CustomNavigationPattern, error) {
	patternObject, err := elem.GetCurrentPattern(CustomNavigationPatternId)

	if err != nil {
		return nil, err
	}

	customNavigationPattern, err := patternObject.QueryInterface(IID_IUIAutomationCustomNavigationPattern)

	if err != nil {
		return nil, err
	}

	return (*CustomNavigationPattern)(unsafe.Pointer(customNavigationPattern)), nil
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
