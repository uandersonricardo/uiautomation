package uiautomation

import (
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
	panic("Not implemented")
}

func (c *AndCondition) GetChildren() ([]*Condition, error) {
	panic("Not implemented")
}

func (c *AndCondition) GetChildrenAsNativeArray() ([]*Condition, error) {
	panic("Not implemented")
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
	panic("Not implemented")
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
	panic("Not implemented")
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
	panic("Not implemented")
}

func (c *PropertyCondition) PropertyValue() (interface{}, error) {
	panic("Not implemented")
}

func (c *PropertyCondition) PropertyConditionFlags() (PropertyConditionFlags, error) {
	panic("Not implemented")
}
