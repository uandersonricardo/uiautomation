package main

import (
	"fmt"

	auto "github.com/uandersonricardo/uiautomation"
)

func focusedElementHandler(el *auto.Element) error {
	// Get the name of the focused element
	fmt.Printf("Focused Element: %v\n", el)

	name, err := el.CurrentName()
	if err != nil {
		return fmt.Errorf("failed to get focused element name: %w", err)
	}

	// Print the name of the focused element
	fmt.Printf("Focused Element Name: %s\n", name)
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

	// Create a new event handler for focused elements
	handler := auto.NewFocusChangedEventHandler(focusedElementHandler)
	defer handler.Release()

	err = uiautomation.AddFocusChangedEventHandler(nil, handler)

	if err != nil {
		fmt.Println("Failed to add focus changed event handler:", err)
		return
	}

	defer uiautomation.RemoveFocusChangedEventHandler(handler)

	// Keep the main thread alive
	done := make(chan bool)
	go func() {
		fmt.Scanln()
		done <- true
	}()

	for {
		select {
		case <-done:
			fmt.Println("Stopping focus monitoring...")
			return
		default:
		}
	}
}
