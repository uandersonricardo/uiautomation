package uiautomation

import (
	"syscall"
	"unsafe"

	"github.com/go-ole/go-ole"
)

type EventId uintptr

const (
	ToolTipOpenedEventId                            EventId = 20000
	ToolTipClosedEventId                            EventId = 20001
	StructureChangedEventId                         EventId = 20002
	MenuOpenedEventId                               EventId = 20003
	AutomationPropertyChangedEventId                EventId = 20004
	AutomationFocusChangedEventId                   EventId = 20005
	AsyncContentLoadedEventId                       EventId = 20006
	MenuClosedEventId                               EventId = 20007
	LayoutInvalidatedEventId                        EventId = 20008
	InvokeInvokedEventId                            EventId = 20009
	SelectionItemElementAddedToSelectionEventId     EventId = 20010
	SelectionItemElementRemovedFromSelectionEventId EventId = 20011
	SelectionItemElementSelectedEventId             EventId = 20012
	SelectionInvalidatedEventId                     EventId = 20013
	TextTextSelectionChangedEventId                 EventId = 20014
	TextTextChangedEventId                          EventId = 20015
	WindowWindowOpenedEventId                       EventId = 20016
	WindowWindowClosedEventId                       EventId = 20017
	MenuModeStartEventId                            EventId = 20018
	MenuModeEndEventId                              EventId = 20019
	InputReachedTargetEventId                       EventId = 20020
	InputReachedOtherElementEventId                 EventId = 20021
	InputDiscardedEventId                           EventId = 20022
	SystemAlertEventId                              EventId = 20023
	LiveRegionChangedEventId                        EventId = 20024
	HostedFragmentRootsInvalidatedEventId           EventId = 20025
	DragDragStartEventId                            EventId = 20026
	DragDragCancelEventId                           EventId = 20027
	DragDragCompleteEventId                         EventId = 20028
	DropTargetDragEnterEventId                      EventId = 20029
	DropTargetDragLeaveEventId                      EventId = 20030
	DropTargetDroppedEventId                        EventId = 20031
	TextEditTextChangedEventId                      EventId = 20032
	TextEditConversionTargetChangedEventId          EventId = 20033
	ChangesEventId                                  EventId = 20034
	NotificationEventId                             EventId = 20035
	ActiveTextPositionChangedEventId                EventId = 20036
)

type EventHandler struct {
	ole.IUnknown
	refCount uintptr
	callback func(*Element, EventId) error
}

type EventHandlerVtbl struct {
	ole.IUnknownVtbl
	HandleAutomationEvent uintptr
}

func NewEventHandler(callback func(*Element, EventId) error) *EventHandler {
	h := &EventHandler{
		IUnknown: ole.IUnknown{},
		refCount: 1,
		callback: callback,
	}

	vtbl := &EventHandlerVtbl{
		IUnknownVtbl: ole.IUnknownVtbl{
			QueryInterface: syscall.NewCallback(h.queryInterface),
			AddRef:         syscall.NewCallback(h.addRef),
			Release:        syscall.NewCallback(h.release),
		},
		HandleAutomationEvent: syscall.NewCallback(h.handleAutomationEvent),
	}

	h.RawVTable = (*interface{})(unsafe.Pointer(vtbl))

	return h
}

func (h *EventHandler) VTable() *EventHandlerVtbl {
	return (*EventHandlerVtbl)(unsafe.Pointer(h.RawVTable))
}

func (h *EventHandler) queryInterface(this *ole.IUnknown, iid *ole.GUID, punk **ole.IUnknown) uintptr {
	if iid == nil || punk == nil {
		return uintptr(ole.E_POINTER)
	}

	*punk = nil

	if ole.IsEqualGUID(iid, ole.IID_IUnknown) ||
		ole.IsEqualGUID(iid, ole.IID_IDispatch) {
		h.addRef(this)
		*punk = this
		return uintptr(ole.S_OK)
	}

	if ole.IsEqualGUID(iid, IID_IUIAutomationEventHandler) {
		h.addRef(this)
		*punk = this
		return uintptr(ole.S_OK)
	}

	return uintptr(ole.E_NOINTERFACE)
}

func (h *EventHandler) addRef(this *ole.IUnknown) uintptr {
	handler := (*EventHandler)(unsafe.Pointer(this))
	handler.refCount++
	return handler.refCount
}

func (h *EventHandler) release(this *ole.IUnknown) uintptr {
	handler := (*EventHandler)(unsafe.Pointer(this))
	handler.refCount--
	return handler.refCount
}

func (h *EventHandler) handleAutomationEvent(this *EventHandler, element *Element, eventId EventId) uintptr {
	err := h.callback(element, eventId)

	if err != nil {
		return uintptr(ole.E_FAIL)
	}

	return uintptr(ole.S_OK)
}

type PropertyChangedEventHandler struct {
	ole.IUnknown
	refCount uintptr
	callback func(*Element, PropertyId, *ole.VARIANT) error
}

type PropertyChangedEventHandlerVtbl struct {
	ole.IUnknownVtbl
	HandlePropertyChangedEvent uintptr
}

func NewPropertyChangedEventHandler(callback func(*Element, PropertyId, *ole.VARIANT) error) *PropertyChangedEventHandler {
	h := &PropertyChangedEventHandler{
		IUnknown: ole.IUnknown{},
		refCount: 1,
		callback: callback,
	}

	vtbl := &PropertyChangedEventHandlerVtbl{
		IUnknownVtbl: ole.IUnknownVtbl{
			QueryInterface: syscall.NewCallback(h.queryInterface),
			AddRef:         syscall.NewCallback(h.addRef),
			Release:        syscall.NewCallback(h.release),
		},
		HandlePropertyChangedEvent: syscall.NewCallback(h.handlePropertyChangedEvent),
	}

	h.RawVTable = (*interface{})(unsafe.Pointer(vtbl))

	return h
}

func (h *PropertyChangedEventHandler) VTable() *PropertyChangedEventHandlerVtbl {
	return (*PropertyChangedEventHandlerVtbl)(unsafe.Pointer(h.RawVTable))
}

func (h *PropertyChangedEventHandler) queryInterface(this *ole.IUnknown, iid *ole.GUID, punk **ole.IUnknown) uintptr {
	if iid == nil || punk == nil {
		return uintptr(ole.E_POINTER)
	}

	*punk = nil

	if ole.IsEqualGUID(iid, ole.IID_IUnknown) ||
		ole.IsEqualGUID(iid, ole.IID_IDispatch) {
		h.addRef(this)
		*punk = this
		return uintptr(ole.S_OK)
	}

	if ole.IsEqualGUID(iid, IID_IUIAutomationPropertyChangedEventHandler) {
		h.addRef(this)
		*punk = this
		return uintptr(ole.S_OK)
	}

	return uintptr(ole.E_NOINTERFACE)
}

func (h *PropertyChangedEventHandler) addRef(this *ole.IUnknown) uintptr {
	handler := (*PropertyChangedEventHandler)(unsafe.Pointer(this))
	handler.refCount++
	return handler.refCount
}

func (h *PropertyChangedEventHandler) release(this *ole.IUnknown) uintptr {
	handler := (*PropertyChangedEventHandler)(unsafe.Pointer(this))
	handler.refCount--
	return handler.refCount
}

func (h *PropertyChangedEventHandler) handlePropertyChangedEvent(this *PropertyChangedEventHandler, element *Element, propertyId PropertyId, newValue *ole.VARIANT) uintptr {
	err := h.callback(element, propertyId, newValue)

	if err != nil {
		return uintptr(ole.E_FAIL)
	}

	return uintptr(ole.S_OK)
}

type FocusChangedEventHandler struct {
	ole.IUnknown
	refCount uintptr
	callback func(*Element) error
}

type FocusChangedEventHandlerVtbl struct {
	ole.IUnknownVtbl
	HandleFocusChangedEvent uintptr
}

func NewFocusChangedEventHandler(callback func(*Element) error) *FocusChangedEventHandler {
	h := &FocusChangedEventHandler{
		IUnknown: ole.IUnknown{},
		refCount: 1,
		callback: callback,
	}

	vtbl := &FocusChangedEventHandlerVtbl{
		IUnknownVtbl: ole.IUnknownVtbl{
			QueryInterface: syscall.NewCallback(h.queryInterface),
			AddRef:         syscall.NewCallback(h.addRef),
			Release:        syscall.NewCallback(h.release),
		},
		HandleFocusChangedEvent: syscall.NewCallback(h.handleFocusChangedEvent),
	}

	h.RawVTable = (*interface{})(unsafe.Pointer(vtbl))

	return h
}

func (h *FocusChangedEventHandler) VTable() *FocusChangedEventHandlerVtbl {
	return (*FocusChangedEventHandlerVtbl)(unsafe.Pointer(h.RawVTable))
}

func (h *FocusChangedEventHandler) queryInterface(this *ole.IUnknown, iid *ole.GUID, punk **ole.IUnknown) uintptr {
	if iid == nil || punk == nil {
		return uintptr(ole.E_POINTER)
	}

	*punk = nil

	if ole.IsEqualGUID(iid, ole.IID_IUnknown) ||
		ole.IsEqualGUID(iid, ole.IID_IDispatch) {
		h.addRef(this)
		*punk = this
		return uintptr(ole.S_OK)
	}

	if ole.IsEqualGUID(iid, IID_IUIAutomationFocusChangedEventHandler) {
		h.addRef(this)
		*punk = this
		return uintptr(ole.S_OK)
	}

	return uintptr(ole.E_NOINTERFACE)
}

func (h *FocusChangedEventHandler) addRef(this *ole.IUnknown) uintptr {
	handler := (*FocusChangedEventHandler)(unsafe.Pointer(this))
	handler.refCount++
	return handler.refCount
}

func (h *FocusChangedEventHandler) release(this *ole.IUnknown) uintptr {
	handler := (*FocusChangedEventHandler)(unsafe.Pointer(this))
	handler.refCount--
	return handler.refCount
}

func (h *FocusChangedEventHandler) handleFocusChangedEvent(this *FocusChangedEventHandler, element *Element) uintptr {
	err := h.callback(element)

	if err != nil {
		return uintptr(ole.E_FAIL)
	}

	return uintptr(ole.S_OK)
}

type StructureChangedEventHandler struct {
	ole.IUnknown
	refCount uintptr
	callback func(*Element, StructureChangeType, *ole.SafeArray) error
}

type StructureChangedEventHandlerVtbl struct {
	ole.IUnknownVtbl
	HandleStructureChangedEvent uintptr
}

func NewStructureChangedEventHandler(callback func(*Element, StructureChangeType, *ole.SafeArray) error) *StructureChangedEventHandler {
	h := &StructureChangedEventHandler{
		IUnknown: ole.IUnknown{},
		refCount: 1,
		callback: callback,
	}

	vtbl := &StructureChangedEventHandlerVtbl{
		IUnknownVtbl: ole.IUnknownVtbl{
			QueryInterface: syscall.NewCallback(h.queryInterface),
			AddRef:         syscall.NewCallback(h.addRef),
			Release:        syscall.NewCallback(h.release),
		},
		HandleStructureChangedEvent: syscall.NewCallback(h.handleStructureChangedEvent),
	}

	h.RawVTable = (*interface{})(unsafe.Pointer(vtbl))

	return h
}

func (h *StructureChangedEventHandler) VTable() *StructureChangedEventHandlerVtbl {
	return (*StructureChangedEventHandlerVtbl)(unsafe.Pointer(h.RawVTable))
}

func (h *StructureChangedEventHandler) queryInterface(this *ole.IUnknown, iid *ole.GUID, punk **ole.IUnknown) uintptr {
	if iid == nil || punk == nil {
		return uintptr(ole.E_POINTER)
	}

	*punk = nil

	if ole.IsEqualGUID(iid, ole.IID_IUnknown) ||
		ole.IsEqualGUID(iid, ole.IID_IDispatch) {
		h.addRef(this)
		*punk = this
		return uintptr(ole.S_OK)
	}

	if ole.IsEqualGUID(iid, IID_IUIAutomationStructureChangedEventHandler) {
		h.addRef(this)
		*punk = this
		return uintptr(ole.S_OK)
	}

	return uintptr(ole.E_NOINTERFACE)
}

func (h *StructureChangedEventHandler) addRef(this *ole.IUnknown) uintptr {
	handler := (*StructureChangedEventHandler)(unsafe.Pointer(this))
	handler.refCount++
	return handler.refCount
}

func (h *StructureChangedEventHandler) release(this *ole.IUnknown) uintptr {
	handler := (*StructureChangedEventHandler)(unsafe.Pointer(this))
	handler.refCount--
	return handler.refCount
}

func (h *StructureChangedEventHandler) handleStructureChangedEvent(this *StructureChangedEventHandler, element *Element, changeType StructureChangeType, runtimeId *ole.SafeArray) uintptr {
	err := h.callback(element, changeType, runtimeId)

	if err != nil {
		return uintptr(ole.E_FAIL)
	}

	return uintptr(ole.S_OK)
}
