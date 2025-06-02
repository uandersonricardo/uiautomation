package main

import (
	"fmt"
	"time"

	auto "github.com/uandersonricardo/uiautomation"
)

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

	timer := *time.NewTicker(1 * time.Second)

	// Wait for a few seconds to allow the user to focus on an element
	for {
		select {
		case <-timer.C:
			// Get focused element
			focusedElement, err := uiautomation.GetFocusedElement()
			if err != nil {
				fmt.Println("Failed to get focused element:", err)
				return
			}

			// Get the name of the focused element
			name, err := focusedElement.CurrentName()
			if err != nil {
				fmt.Println("Failed to get focused element name:", err)
				return
			}

			// Print the name of the focused element
			fmt.Printf("Focused Element Name: %s\n", name)

			variant, err := focusedElement.GetCurrentPropertyValue(auto.RuntimeIdPropertyId)
			if err != nil {
				fmt.Println("Failed to get focused element RuntimeId:", err)
				return
			}

			// Print variant
			fmt.Printf("Focused Element RuntimeId: %s\n", variant)

			pattern, err := focusedElement.GetCurrentPattern(auto.TextPatternId)
			if err != nil {
				fmt.Println("Failed to get focused element TextPattern:", err)
				return
			}

			// Print pattern
			fmt.Printf("Focused Element TextPattern: %s\n", pattern)

			runtimeId, err := focusedElement.GetRuntimeId()
			if err != nil {
				fmt.Println("Failed to get focused element RuntimeId:", err)
				return
			}

			// Print the runtime ID
			fmt.Printf("Focused Element RuntimeId: %v\n", runtimeId)
		default:
			// Do nothing, just wait for the user to focus on an element
		}
	}
}
