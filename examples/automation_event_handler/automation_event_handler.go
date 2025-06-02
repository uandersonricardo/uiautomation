package main

import (
	"fmt"

	auto "github.com/uandersonricardo/uiautomation"
)

func handleAutomationEvent(element *auto.Element, eventId auto.EventId) error {
	switch eventId {
	case auto.ToolTipOpenedEventId:
		fmt.Println("Tooltip opened")
	case auto.ToolTipClosedEventId:
		fmt.Println("Tooltip closed")
	case auto.WindowWindowOpenedEventId:
		fmt.Println("Window opened")
	case auto.WindowWindowClosedEventId:
		fmt.Println("Window closed")
	default:
		fmt.Printf("Unhandled event: %v\n", eventId)
	}

	return nil
}

func main() {
	// Initialize COM library
	if err := auto.CoInitialize(); err != nil {
		fmt.Println("Failed to initialize COM library:", err)
		return
	}
	defer auto.CoUninitialize()

	// Create a new UI Automation object
	uiautomation, err := auto.NewUIAutomation()
	if err != nil {
		fmt.Println("Failed to create UI Automation object:", err)
		return
	}
	defer uiautomation.Release()

	// Get root element
	root, err := uiautomation.GetRootElement()
	if err != nil {
		fmt.Println("Failed to get root element:", err)
		return
	}

	fmt.Println("Got root element:", root)

	// Create a event handler
	handler := auto.NewEventHandler(handleAutomationEvent)
	defer handler.Release()

	// Add ToolTipOpenedEvent handler
	uiautomation.AddAutomationEventHandler(auto.ToolTipOpenedEventId, root, auto.TreeScopeSubtree, nil, handler)
	defer uiautomation.RemoveAutomationEventHandler(auto.ToolTipOpenedEventId, root, handler)

	// Add ToolTipClosedEvent handler
	uiautomation.AddAutomationEventHandler(auto.ToolTipClosedEventId, root, auto.TreeScopeSubtree, nil, handler)
	defer uiautomation.RemoveAutomationEventHandler(auto.ToolTipClosedEventId, root, handler)

	// Add WindowOpenedEvent handler
	uiautomation.AddAutomationEventHandler(auto.WindowWindowOpenedEventId, root, auto.TreeScopeSubtree, nil, handler)
	defer uiautomation.RemoveAutomationEventHandler(auto.WindowWindowOpenedEventId, root, handler)

	// Add WindowClosedEvent handler
	uiautomation.AddAutomationEventHandler(auto.WindowWindowClosedEventId, root, auto.TreeScopeSubtree, nil, handler)
	defer uiautomation.RemoveAutomationEventHandler(auto.WindowWindowClosedEventId, root, handler)

	// Keep the main thread alive
	done := make(chan bool)
	go func() {
		fmt.Scanln()
		done <- true
	}()

	for {
		select {
		case <-done:
			fmt.Println("Stopping event monitoring...")
			return
		default:
		}
	}
}
