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

			// Get the bounding rectangle of the focused element
			boundingRect, err := focusedElement.CurrentBoundingRectangle()
			if err != nil {
				fmt.Println("Failed to get bounding rectangle:", err)
				return
			}

			// Print the bounding rectangle
			fmt.Printf("Bounding Rectangle: %+v\n", boundingRect)
		default:
			// Do nothing, just wait for the user to focus on an element
		}
	}
}
