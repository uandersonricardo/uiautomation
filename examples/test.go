package main

import (
	"fmt"

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

	// Get the root element of the desktop
	rootElement, err := uiautomation.GetRootElement()
	if err != nil {
		fmt.Println("Failed to get root element:", err)
		return
	}
	defer rootElement.Release()

	fmt.Println("Root element obtained successfully")

	// Get the name of the root element
	name, err := rootElement.CurrentName()
	if err != nil {
		fmt.Println("Failed to get root element name:", err)
		return
	}
	fmt.Println("Root element name:", name)
}
