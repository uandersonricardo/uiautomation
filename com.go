package uiautomation

import "github.com/go-ole/go-ole"

func CoInitialize() error {
	return ole.CoInitializeEx(0, ole.COINIT_MULTITHREADED)
}

func CoUninitialize() {
	ole.CoUninitialize()
}
