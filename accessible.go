package uiautomation

import (
	"unsafe"

	"github.com/go-ole/go-ole"
)

type Accessible struct {
	ole.IDispatch
}

type AccessibleVtbl struct {
	ole.IDispatchVtbl
	Get_accParent           uintptr
	Get_accChildCount       uintptr
	Get_accChild            uintptr
	Get_accName             uintptr
	Get_accValue            uintptr
	Get_accDescription      uintptr
	Get_accRole             uintptr
	Get_accState            uintptr
	Get_accHelp             uintptr
	Get_accHelpTopic        uintptr
	Get_accKeyboardShortcut uintptr
	Get_accFocus            uintptr
	Get_accSelection        uintptr
	Get_accDefaultAction    uintptr
	AccSelect               uintptr
	AccLocation             uintptr
	AccNavigate             uintptr
	AccHitTest              uintptr
	AccDoDefaultAction      uintptr
	Put_accName             uintptr
	Put_accValue            uintptr
}

func (ac *Accessible) VTable() *AccessibleVtbl {
	return (*AccessibleVtbl)(unsafe.Pointer(ac.RawVTable))
}

type UIAutomationLegacyAccessiblePattern struct {
	ole.IUnknown
}

type UIAutomationLegacyAccessiblePatternVtbl struct {
	ole.IUnknownVtbl
	DoDefaultAction             uintptr
	Get_CachedChildId           uintptr
	Get_CachedDefaultAction     uintptr
	Get_CachedDescription       uintptr
	Get_CachedHelp              uintptr
	Get_CachedKeyboardShortcut  uintptr
	Get_CachedName              uintptr
	Get_CachedRole              uintptr
	Get_CachedState             uintptr
	Get_CachedValue             uintptr
	Get_CurrentChildId          uintptr
	Get_CurrentDefaultAction    uintptr
	Get_CurrentDescription      uintptr
	Get_CurrentHelp             uintptr
	Get_CurrentKeyboardShortcut uintptr
	Get_CurrentName             uintptr
	Get_CurrentRole             uintptr
	Get_CurrentState            uintptr
	Get_CurrentValue            uintptr
	GetCachedSelection          uintptr
	GetCurrentSelection         uintptr
	GetAccessible               uintptr
	Select                      uintptr
	SetValue                    uintptr
}

func (lap *UIAutomationLegacyAccessiblePattern) VTable() *UIAutomationLegacyAccessiblePatternVtbl {
	return (*UIAutomationLegacyAccessiblePatternVtbl)(unsafe.Pointer(lap.RawVTable))
}
