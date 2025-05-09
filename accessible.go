package uiautomation

import (
	"syscall"
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

func (ac *Accessible) Parent() (*Accessible, error) {
	var parent *Accessible

	hr, _, _ := syscall.SyscallN(
		ac.VTable().Get_accParent,
		uintptr(unsafe.Pointer(ac)),
		uintptr(unsafe.Pointer(&parent)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return parent, nil
}

func (ac *Accessible) ChildCount() (int32, error) {
	var count int32

	hr, _, _ := syscall.SyscallN(
		ac.VTable().Get_accChildCount,
		uintptr(unsafe.Pointer(ac)),
		uintptr(unsafe.Pointer(&count)),
	)

	if hr != 0 {
		return 0, ole.NewError(hr)
	}

	return count, nil
}

func (ac *Accessible) Child(index int32) (*Accessible, error) {
	var child *Accessible

	hr, _, _ := syscall.SyscallN(
		ac.VTable().Get_accChild,
		uintptr(unsafe.Pointer(ac)),
		uintptr(index),
		uintptr(unsafe.Pointer(&child)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return child, nil
}

func (ac *Accessible) Name() (string, error) {
	var name *uint16

	hr, _, _ := syscall.SyscallN(
		ac.VTable().Get_accName,
		uintptr(unsafe.Pointer(ac)),
		uintptr(unsafe.Pointer(&name)),
	)

	if hr != 0 {
		return "", ole.NewError(hr)
	}

	return ole.BstrToString(name), nil
}

func (ac *Accessible) Value() (string, error) {
	var value *uint16

	hr, _, _ := syscall.SyscallN(
		ac.VTable().Get_accValue,
		uintptr(unsafe.Pointer(ac)),
		uintptr(unsafe.Pointer(&value)),
	)

	if hr != 0 {
		return "", ole.NewError(hr)
	}

	return ole.BstrToString(value), nil
}

func (ac *Accessible) Description() (string, error) {
	var description *uint16

	hr, _, _ := syscall.SyscallN(
		ac.VTable().Get_accDescription,
		uintptr(unsafe.Pointer(ac)),
		uintptr(unsafe.Pointer(&description)),
	)

	if hr != 0 {
		return "", ole.NewError(hr)
	}

	return ole.BstrToString(description), nil
}

func (ac *Accessible) Role() (int32, error) {
	var role int32

	hr, _, _ := syscall.SyscallN(
		ac.VTable().Get_accRole,
		uintptr(unsafe.Pointer(ac)),
		uintptr(unsafe.Pointer(&role)),
	)

	if hr != 0 {
		return 0, ole.NewError(hr)
	}

	return role, nil
}

func (ac *Accessible) State() (int32, error) {
	var state int32

	hr, _, _ := syscall.SyscallN(
		ac.VTable().Get_accState,
		uintptr(unsafe.Pointer(ac)),
		uintptr(unsafe.Pointer(&state)),
	)

	if hr != 0 {
		return 0, ole.NewError(hr)
	}

	return state, nil
}

func (ac *Accessible) Help() (string, error) {
	var help *uint16

	hr, _, _ := syscall.SyscallN(
		ac.VTable().Get_accHelp,
		uintptr(unsafe.Pointer(ac)),
		uintptr(unsafe.Pointer(&help)),
	)

	if hr != 0 {
		return "", ole.NewError(hr)
	}

	return ole.BstrToString(help), nil
}

func (ac *Accessible) HelpTopic() (string, error) {
	var topic *uint16

	hr, _, _ := syscall.SyscallN(
		ac.VTable().Get_accHelpTopic,
		uintptr(unsafe.Pointer(ac)),
		uintptr(unsafe.Pointer(&topic)),
	)

	if hr != 0 {
		return "", ole.NewError(hr)
	}

	return ole.BstrToString(topic), nil
}

func (ac *Accessible) KeyboardShortcut() (string, error) {
	var shortcut *uint16

	hr, _, _ := syscall.SyscallN(
		ac.VTable().Get_accKeyboardShortcut,
		uintptr(unsafe.Pointer(ac)),
		uintptr(unsafe.Pointer(&shortcut)),
	)

	if hr != 0 {
		return "", ole.NewError(hr)
	}

	return ole.BstrToString(shortcut), nil
}

func (ac *Accessible) Focus() (int32, error) {
	var focus int32

	hr, _, _ := syscall.SyscallN(
		ac.VTable().Get_accFocus,
		uintptr(unsafe.Pointer(ac)),
		uintptr(unsafe.Pointer(&focus)),
	)

	if hr != 0 {
		return 0, ole.NewError(hr)
	}

	return focus, nil
}

func (ac *Accessible) Selection() (int32, error) {
	var selection int32

	hr, _, _ := syscall.SyscallN(
		ac.VTable().Get_accSelection,
		uintptr(unsafe.Pointer(ac)),
		uintptr(unsafe.Pointer(&selection)),
	)

	if hr != 0 {
		return 0, ole.NewError(hr)
	}

	return selection, nil
}

func (ac *Accessible) DefaultAction() (string, error) {
	var action *uint16

	hr, _, _ := syscall.SyscallN(
		ac.VTable().Get_accDefaultAction,
		uintptr(unsafe.Pointer(ac)),
		uintptr(unsafe.Pointer(&action)),
	)

	if hr != 0 {
		return "", ole.NewError(hr)
	}

	return ole.BstrToString(action), nil
}

func (ac *Accessible) Select(flags int32) error {
	hr, _, _ := syscall.SyscallN(
		ac.VTable().AccSelect,
		uintptr(unsafe.Pointer(ac)),
		uintptr(flags),
	)

	if hr != 0 {
		return ole.NewError(hr)
	}

	return nil
}

func (ac *Accessible) Location() (int32, int32, int32, int32, error) {
	var left, top, width, height int32

	hr, _, _ := syscall.SyscallN(
		ac.VTable().AccLocation,
		uintptr(unsafe.Pointer(ac)),
		uintptr(unsafe.Pointer(&left)),
		uintptr(unsafe.Pointer(&top)),
		uintptr(unsafe.Pointer(&width)),
		uintptr(unsafe.Pointer(&height)),
	)

	if hr != 0 {
		return 0, 0, 0, 0, ole.NewError(hr)
	}

	return left, top, width, height, nil
}

func (ac *Accessible) Navigate(navDir int32) (*Accessible, error) {
	var child *Accessible

	hr, _, _ := syscall.SyscallN(
		ac.VTable().AccNavigate,
		uintptr(unsafe.Pointer(ac)),
		uintptr(navDir),
		uintptr(unsafe.Pointer(&child)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return child, nil
}

func (ac *Accessible) HitTest(x, y int32) (*Accessible, error) {
	var child *Accessible

	hr, _, _ := syscall.SyscallN(
		ac.VTable().AccHitTest,
		uintptr(unsafe.Pointer(ac)),
		uintptr(x),
		uintptr(y),
		uintptr(unsafe.Pointer(&child)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return child, nil
}

func (ac *Accessible) DoDefaultAction() error {
	hr, _, _ := syscall.SyscallN(
		ac.VTable().AccDoDefaultAction,
		uintptr(unsafe.Pointer(ac)),
	)

	if hr != 0 {
		return ole.NewError(hr)
	}

	return nil
}

func (ac *Accessible) SetName(name string) error {
	namePtr, err := syscall.UTF16PtrFromString(name)

	if err != nil {
		return err
	}

	hr, _, _ := syscall.SyscallN(
		ac.VTable().Put_accName,
		uintptr(unsafe.Pointer(ac)),
		uintptr(unsafe.Pointer(namePtr)),
	)

	if hr != 0 {
		return ole.NewError(hr)
	}

	return nil
}

func (ac *Accessible) SetValue(value string) error {
	valuePtr, err := syscall.UTF16PtrFromString(value)

	if err != nil {
		return err
	}

	hr, _, _ := syscall.SyscallN(
		ac.VTable().Put_accValue,
		uintptr(unsafe.Pointer(ac)),
		uintptr(unsafe.Pointer(valuePtr)),
	)

	if hr != 0 {
		return ole.NewError(hr)
	}

	return nil
}
