package uiautomation

import (
	"unsafe"

	"github.com/go-ole/go-ole"
)

type UIAutomationCondition struct {
	ole.IUnknown
}

type UIAutomationConditionVtbl struct {
	ole.IUnknownVtbl
}

type UIAutomationAndCondition struct {
	UIAutomationCondition
}

type UIAutomationAndConditionVtbl struct {
	UIAutomationConditionVtbl
	Get_ChildCount           uintptr
	GetChildrenAsNativeArray uintptr
	GetChildren              uintptr
}

func (c *UIAutomationAndCondition) VTable() *UIAutomationAndConditionVtbl {
	return (*UIAutomationAndConditionVtbl)(unsafe.Pointer(c.RawVTable))
}

func (c *UIAutomationAndCondition) ChildCount() (int32, error) {
	panic("Not implemented")
}

func (c *UIAutomationAndCondition) GetChildren() ([]*UIAutomationCondition, error) {
	panic("Not implemented")
}

func (c *UIAutomationAndCondition) GetChildrenAsNativeArray() ([]*UIAutomationCondition, error) {
	panic("Not implemented")
}

type UIAutomationBoolCondition struct {
	UIAutomationCondition
}

type UIAutomationBoolConditionVtbl struct {
	UIAutomationConditionVtbl
	Get_BooleanValue uintptr
}

func (c *UIAutomationBoolCondition) VTable() *UIAutomationBoolConditionVtbl {
	return (*UIAutomationBoolConditionVtbl)(unsafe.Pointer(c.RawVTable))
}

func (c *UIAutomationBoolCondition) BooleanValue() (bool, error) {
	panic("Not implemented")
}

type UIAutomationNotCondition struct {
	UIAutomationCondition
}

type UIAutomationNotConditionVtbl struct {
	UIAutomationConditionVtbl
	GetChild uintptr
}

func (c *UIAutomationNotCondition) VTable() *UIAutomationNotConditionVtbl {
	return (*UIAutomationNotConditionVtbl)(unsafe.Pointer(c.RawVTable))
}

func (c *UIAutomationNotCondition) GetChild() (*UIAutomationCondition, error) {
	panic("Not implemented")
}

type UIAutomationPropertyCondition struct {
	UIAutomationCondition
}

type UIAutomationPropertyConditionVtbl struct {
	UIAutomationConditionVtbl
	Get_PropertyId             uintptr
	Get_PropertyValue          uintptr
	Get_PropertyConditionFlags uintptr
}

func (c *UIAutomationPropertyCondition) VTable() *UIAutomationPropertyConditionVtbl {
	return (*UIAutomationPropertyConditionVtbl)(unsafe.Pointer(c.RawVTable))
}

func (c *UIAutomationPropertyCondition) PropertyId() (PropertyId, error) {
	panic("Not implemented")
}

func (c *UIAutomationPropertyCondition) PropertyValue() (interface{}, error) {
	panic("Not implemented")
}

func (c *UIAutomationPropertyCondition) PropertyConditionFlags() (PropertyConditionFlags, error) {
	panic("Not implemented")
}
