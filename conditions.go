package uiautomation

import (
	"syscall"
	"unsafe"

	"github.com/go-ole/go-ole"
)

type Condition struct {
	ole.IUnknown
}

type ConditionVtbl struct {
	ole.IUnknownVtbl
}

type AndCondition struct {
	Condition
}

type AndConditionVtbl struct {
	ConditionVtbl
	Get_ChildCount           uintptr
	GetChildrenAsNativeArray uintptr
	GetChildren              uintptr
}

func (c *AndCondition) VTable() *AndConditionVtbl {
	return (*AndConditionVtbl)(unsafe.Pointer(c.RawVTable))
}

func (c *AndCondition) ChildCount() (int32, error) {
	var count int32

	hr, _, _ := syscall.SyscallN(
		c.VTable().Get_ChildCount,
		uintptr(unsafe.Pointer(c)),
		uintptr(unsafe.Pointer(&count)),
	)

	if hr != 0 {
		return 0, ole.NewError(hr)
	}

	return count, nil
}

func (c *AndCondition) GetChildren() (*ole.SafeArray, error) {
	var children *ole.SafeArray

	hr, _, _ := syscall.SyscallN(
		c.VTable().GetChildren,
		uintptr(unsafe.Pointer(c)),
		uintptr(unsafe.Pointer(&children)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return children, nil
}

func (c *AndCondition) GetChildrenAsNativeArray() ([]*Condition, error) {
	var childrenPtr **Condition
	var count int32

	hr, _, _ := syscall.SyscallN(
		c.VTable().GetChildrenAsNativeArray,
		uintptr(unsafe.Pointer(c)),
		uintptr(unsafe.Pointer(&childrenPtr)),
		uintptr(unsafe.Pointer(&count)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	slice := (*[1 << 30]*Condition)(unsafe.Pointer(childrenPtr))[:count:count]
	children := make([]*Condition, count)
	copy(children, slice)

	ole.CoTaskMemFree(uintptr(unsafe.Pointer(childrenPtr)))

	return children, nil
}

type BoolCondition struct {
	Condition
}

type BoolConditionVtbl struct {
	ConditionVtbl
	Get_BooleanValue uintptr
}

func (c *BoolCondition) VTable() *BoolConditionVtbl {
	return (*BoolConditionVtbl)(unsafe.Pointer(c.RawVTable))
}

func (c *BoolCondition) BooleanValue() (bool, error) {
	var value bool

	hr, _, _ := syscall.SyscallN(
		c.VTable().Get_BooleanValue,
		uintptr(unsafe.Pointer(c)),
		uintptr(unsafe.Pointer(&value)),
	)

	if hr != 0 {
		return false, ole.NewError(hr)
	}

	return value, nil
}

type NotCondition struct {
	Condition
}

type NotConditionVtbl struct {
	ConditionVtbl
	GetChild uintptr
}

func (c *NotCondition) VTable() *NotConditionVtbl {
	return (*NotConditionVtbl)(unsafe.Pointer(c.RawVTable))
}

func (c *NotCondition) GetChild() (*Condition, error) {
	var child *Condition

	hr, _, _ := syscall.SyscallN(
		c.VTable().GetChild,
		uintptr(unsafe.Pointer(c)),
		uintptr(unsafe.Pointer(&child)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return child, nil
}

type PropertyCondition struct {
	Condition
}

type PropertyConditionVtbl struct {
	ConditionVtbl
	Get_PropertyId             uintptr
	Get_PropertyValue          uintptr
	Get_PropertyConditionFlags uintptr
}

func (c *PropertyCondition) VTable() *PropertyConditionVtbl {
	return (*PropertyConditionVtbl)(unsafe.Pointer(c.RawVTable))
}

func (c *PropertyCondition) PropertyId() (PropertyId, error) {
	var propertyId PropertyId

	hr, _, _ := syscall.SyscallN(
		c.VTable().Get_PropertyId,
		uintptr(unsafe.Pointer(c)),
		uintptr(unsafe.Pointer(&propertyId)),
	)

	if hr != 0 {
		return 0, ole.NewError(hr)
	}

	return propertyId, nil
}

func (c *PropertyCondition) PropertyValue() (*ole.VARIANT, error) {
	var value *ole.VARIANT

	hr, _, _ := syscall.SyscallN(
		c.VTable().Get_PropertyValue,
		uintptr(unsafe.Pointer(c)),
		uintptr(unsafe.Pointer(&value)),
	)

	if hr != 0 {
		return nil, ole.NewError(hr)
	}

	return value, nil
}

func (c *PropertyCondition) PropertyConditionFlags() (PropertyConditionFlags, error) {
	var flags PropertyConditionFlags

	hr, _, _ := syscall.SyscallN(
		c.VTable().Get_PropertyConditionFlags,
		uintptr(unsafe.Pointer(c)),
		uintptr(unsafe.Pointer(&flags)),
	)

	if hr != 0 {
		return 0, ole.NewError(hr)
	}

	return flags, nil
}
