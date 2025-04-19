package uiautomation

import "github.com/go-ole/go-ole"

func CoInitialize() error {
	return ole.CoInitialize(0)
}

func CoUninitialize() {
	ole.CoUninitialize()
}
