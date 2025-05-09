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
}

type EventHandlerVtbl struct {
	ole.IUnknownVtbl
	HandleAutomationEvent uintptr
}

func (eh *EventHandler) VTable() *EventHandlerVtbl {
	return (*EventHandlerVtbl)(unsafe.Pointer(eh.RawVTable))
}

func (eh *EventHandler) HandleAutomationEvent(element *Element, eventId EventId) error {
	hr, _, _ := syscall.SyscallN(
		eh.VTable().HandleAutomationEvent,
		uintptr(unsafe.Pointer(eh)),
		uintptr(unsafe.Pointer(element)),
		uintptr(eventId),
	)

	if hr != 0 {
		return ole.NewError(hr)
	}

	return nil
}

type PropertyChangedEventHandler struct {
	ole.IUnknown
}

type PropertyChangedEventHandlerVtbl struct {
	ole.IUnknownVtbl
	HandlePropertyChangedEvent uintptr
}

func (eh *PropertyChangedEventHandler) VTable() *PropertyChangedEventHandlerVtbl {
	return (*PropertyChangedEventHandlerVtbl)(unsafe.Pointer(eh.RawVTable))
}

func (eh *PropertyChangedEventHandler) HandlePropertyChangedEvent(sender *Element, propertyId int32, newValue interface{}) error {
	hr, _, _ := syscall.SyscallN(
		eh.VTable().HandlePropertyChangedEvent,
		uintptr(unsafe.Pointer(eh)),
		uintptr(unsafe.Pointer(sender)),
		uintptr(propertyId),
		uintptr(unsafe.Pointer(&newValue)),
	)

	if hr != 0 {
		return ole.NewError(hr)
	}

	return nil
}

type FocusChangedEventHandler struct {
	ole.IUnknown
}
type FocusChangedEventHandlerVtbl struct {
	ole.IUnknownVtbl
	HandleFocusChangedEvent uintptr
}

func (eh *FocusChangedEventHandler) VTable() *FocusChangedEventHandlerVtbl {
	return (*FocusChangedEventHandlerVtbl)(unsafe.Pointer(eh.RawVTable))
}

func (eh *FocusChangedEventHandler) HandleFocusChangedEvent(element *Element) error {
	hr, _, _ := syscall.SyscallN(
		eh.VTable().HandleFocusChangedEvent,
		uintptr(unsafe.Pointer(eh)),
		uintptr(unsafe.Pointer(element)),
	)

	if hr != 0 {
		return ole.NewError(hr)
	}

	return nil
}

type StructureChangedEventHandler struct {
	ole.IUnknown
}
type StructureChangedEventHandlerVtbl struct {
	ole.IUnknownVtbl
	HandleStructureChangedEvent uintptr
}

func (eh *StructureChangedEventHandler) VTable() *StructureChangedEventHandlerVtbl {
	return (*StructureChangedEventHandlerVtbl)(unsafe.Pointer(eh.RawVTable))
}

func (eh *StructureChangedEventHandler) HandleStructureChangedEvent(sender *Element, changeType StructureChangeType, runtimeId []int32) error {
	panic("Not implemented")
}
