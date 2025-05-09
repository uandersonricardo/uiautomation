package uiautomation

type ControlTypeId uintptr

const (
	ButtonControlTypeId       ControlTypeId = 50000
	CalendarControlTypeId     ControlTypeId = 50001
	CheckBoxControlTypeId     ControlTypeId = 50002
	ComboBoxControlTypeId     ControlTypeId = 50003
	EditControlTypeId         ControlTypeId = 50004
	HyperlinkControlTypeId    ControlTypeId = 50005
	ImageControlTypeId        ControlTypeId = 50006
	ListItemControlTypeId     ControlTypeId = 50007
	ListControlTypeId         ControlTypeId = 50008
	MenuControlTypeId         ControlTypeId = 50009
	MenuBarControlTypeId      ControlTypeId = 50010
	MenuItemControlTypeId     ControlTypeId = 50011
	ProgressBarControlTypeId  ControlTypeId = 50012
	RadioButtonControlTypeId  ControlTypeId = 50013
	ScrollBarControlTypeId    ControlTypeId = 50014
	SliderControlTypeId       ControlTypeId = 50015
	SpinnerControlTypeId      ControlTypeId = 50016
	StatusBarControlTypeId    ControlTypeId = 50017
	TabControlTypeId          ControlTypeId = 50018
	TabItemControlTypeId      ControlTypeId = 50019
	TextControlTypeId         ControlTypeId = 50020
	ToolBarControlTypeId      ControlTypeId = 50021
	ToolTipControlTypeId      ControlTypeId = 50022
	TreeControlTypeId         ControlTypeId = 50023
	TreeItemControlTypeId     ControlTypeId = 50024
	CustomControlTypeId       ControlTypeId = 50025
	GroupControlTypeId        ControlTypeId = 50026
	ThumbControlTypeId        ControlTypeId = 50027
	DataGridControlTypeId     ControlTypeId = 50028
	DataItemControlTypeId     ControlTypeId = 50029
	DocumentControlTypeId     ControlTypeId = 50030
	SplitButtonControlTypeId  ControlTypeId = 50031
	WindowControlTypeId       ControlTypeId = 50032
	PaneControlTypeId         ControlTypeId = 50033
	HeaderControlTypeId       ControlTypeId = 50034
	HeaderItemControlTypeId   ControlTypeId = 50035
	TableControlTypeId        ControlTypeId = 50036
	TitleBarControlTypeId     ControlTypeId = 50037
	SeparatorControlTypeId    ControlTypeId = 50038
	SemanticZoomControlTypeId ControlTypeId = 50039
	AppBarControlTypeId       ControlTypeId = 50040
)

var ControlTypeNames = map[ControlTypeId]string{
	ButtonControlTypeId:       "ButtonControl",
	CalendarControlTypeId:     "CalendarControl",
	CheckBoxControlTypeId:     "CheckBoxControl",
	ComboBoxControlTypeId:     "ComboBoxControl",
	EditControlTypeId:         "EditControl",
	HyperlinkControlTypeId:    "HyperlinkControl",
	ImageControlTypeId:        "ImageControl",
	ListItemControlTypeId:     "ListItemControl",
	ListControlTypeId:         "ListControl",
	MenuControlTypeId:         "MenuControl",
	MenuBarControlTypeId:      "MenuBarControl",
	MenuItemControlTypeId:     "MenuItemControl",
	ProgressBarControlTypeId:  "ProgressBarControl",
	RadioButtonControlTypeId:  "RadioButtonControl",
	ScrollBarControlTypeId:    "ScrollBarControl",
	SliderControlTypeId:       "SliderControl",
	SpinnerControlTypeId:      "SpinnerControl",
	StatusBarControlTypeId:    "StatusBarControl",
	TabControlTypeId:          "TabControl",
	TabItemControlTypeId:      "TabItemControl",
	TextControlTypeId:         "TextControl",
	ToolBarControlTypeId:      "ToolBarControl",
	ToolTipControlTypeId:      "ToolTipControl",
	TreeControlTypeId:         "TreeControl",
	TreeItemControlTypeId:     "TreeItemControl",
	CustomControlTypeId:       "CustomControl",
	GroupControlTypeId:        "GroupControl",
	ThumbControlTypeId:        "ThumbControl",
	DataGridControlTypeId:     "DataGridControl",
	DataItemControlTypeId:     "DataItemControl",
	DocumentControlTypeId:     "DocumentControl",
	SplitButtonControlTypeId:  "SplitButtonControl",
	WindowControlTypeId:       "WindowControl",
	PaneControlTypeId:         "PaneControl",
	HeaderControlTypeId:       "HeaderControl",
	HeaderItemControlTypeId:   "HeaderItemControl",
	TableControlTypeId:        "TableControl",
	TitleBarControlTypeId:     "TitleBarControl",
	SeparatorControlTypeId:    "SeparatorControl",
	SemanticZoomControlTypeId: "SemanticZoomControl",
	AppBarControlTypeId:       "AppBarControl",
}

func ControlTypeIdFromName(name string) ControlTypeId {
	for id, n := range ControlTypeNames {
		if n == name {
			return id
		}
	}

	return 0
}

func ControlTypeNameFromId(id ControlTypeId) string {
	for i, n := range ControlTypeNames {
		if i == id {
			return n
		}
	}

	return ""
}
