package uiautomation

import "errors"

var (
	ErrElementIsNil    = errors.New("uiautomation: element is nil")
	ErrPatternNotFound = errors.New("uiautomation: pattern not found (likely unsupported by element)")
)
