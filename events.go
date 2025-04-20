package uiautomation

import (
	"syscall"
	"unsafe"

	"github.com/go-ole/go-ole"
)

type EventId uintptr

const (
	ActiveTextPositionChangedEventId                EventId = 20036
	AsyncContentLoadedEventId                       EventId = 20006
	AutomationFocusChangedEventId                   EventId = 20005
	AutomationPropertyChangedEventId                EventId = 20004
	ChangesEventId                                  EventId = 20034
	DragDragCancelEventId                           EventId = 20027
	DragDragCompleteEventId                         EventId = 20028
	DragDragStartEventId                            EventId = 20026
	DropTargetDragEnterEventId                      EventId = 20029
	DropTargetDragLeaveEventId                      EventId = 20030
	DropTargetDroppedEventId                        EventId = 20031
	HostedFragmentRootsInvalidatedEventId           EventId = 20025
	InputDiscardedEventId                           EventId = 20022
	InputReachedOtherElementEventId                 EventId = 20021
	InputReachedTargetEventId                       EventId = 20020
	InvokeInvokedEventId                            EventId = 20009
	LayoutInvalidatedEventId                        EventId = 20008
	LiveRegionChangedEventId                        EventId = 20024
	MenuClosedEventId                               EventId = 20007
	MenuModeEndEventId                              EventId = 20019
	MenuModeStartEventId                            EventId = 20018
	MenuOpenedEventId                               EventId = 20003
	NotificationEventId                             EventId = 20035
	SelectionInvalidatedEventId                     EventId = 20013
	SelectionItemElementAddedToSelectionEventId     EventId = 20010
	SelectionItemElementRemovedFromSelectionEventId EventId = 20011
	SelectionItemElementSelectedEventId             EventId = 20012
	StructureChangedEventId                         EventId = 20002
	SystemAlertEventId                              EventId = 20023
	TextTextChangedEventId                          EventId = 20015
	TextTextSelectionChangedEventId                 EventId = 20014
	TextEditConversionTargetChangedEventId          EventId = 20033
	TextEditTextChangedEventId                      EventId = 20032
	ToolTipClosedEventId                            EventId = 20001
	ToolTipOpenedEventId                            EventId = 20000
	WindowWindowClosedEventId                       EventId = 20017
	WindowWindowOpenedEventId                       EventId = 20016
)

type UIAutomationEventHandler struct {
	ole.IUnknown
}

type UIAutomationEventHandlerVtbl struct {
	ole.IUnknownVtbl
	HandleAutomationEvent uintptr
}

func (eh *UIAutomationEventHandler) VTable() *UIAutomationEventHandlerVtbl {
	return (*UIAutomationEventHandlerVtbl)(unsafe.Pointer(eh.RawVTable))
}

func (eh *UIAutomationEventHandler) HandleAutomationEvent(element *UIAutomationElement, eventId EventId) error {
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

type UIAutomationPropertyChangedEventHandler struct {
	ole.IUnknown
}

type UIAutomationPropertyChangedEventHandlerVtbl struct {
	ole.IUnknownVtbl
	HandlePropertyChangedEvent uintptr
}

func (eh *UIAutomationPropertyChangedEventHandler) VTable() *UIAutomationPropertyChangedEventHandlerVtbl {
	return (*UIAutomationPropertyChangedEventHandlerVtbl)(unsafe.Pointer(eh.RawVTable))
}

func (eh *UIAutomationPropertyChangedEventHandler) HandlePropertyChangedEvent(sender *UIAutomationElement, propertyId int32, newValue interface{}) error {
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

type UIAutomationFocusChangedEventHandler struct {
	ole.IUnknown
}
type UIAutomationFocusChangedEventHandlerVtbl struct {
	ole.IUnknownVtbl
	HandleFocusChangedEvent uintptr
}

func (eh *UIAutomationFocusChangedEventHandler) VTable() *UIAutomationFocusChangedEventHandlerVtbl {
	return (*UIAutomationFocusChangedEventHandlerVtbl)(unsafe.Pointer(eh.RawVTable))
}

func (eh *UIAutomationFocusChangedEventHandler) HandleFocusChangedEvent(element *UIAutomationElement) error {
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

type UIAutomationStructureChangedEventHandler struct {
	ole.IUnknown
}
type UIAutomationStructureChangedEventHandlerVtbl struct {
	ole.IUnknownVtbl
	HandleStructureChangedEvent uintptr
}

func (eh *UIAutomationStructureChangedEventHandler) VTable() *UIAutomationStructureChangedEventHandlerVtbl {
	return (*UIAutomationStructureChangedEventHandlerVtbl)(unsafe.Pointer(eh.RawVTable))
}

func (eh *UIAutomationStructureChangedEventHandler) HandleStructureChangedEvent(sender *UIAutomationElement, changeType StructureChangeType, runtimeId []int32) error {
	panic("Not implemented")
}
