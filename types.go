package uiautomation

type Rect struct {
	Left   uint32
	Top    uint32
	Right  uint32
	Bottom uint32
}

func (r Rect) Width() uint32 {
	return r.Right - r.Left
}

func (r Rect) Height() uint32 {
	return r.Bottom - r.Top
}

type Point struct {
	X int32
	Y int32
}
