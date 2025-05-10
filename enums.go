package uiautomation

type PropertyConditionFlags uint32

const (
	PropertyConditionFlagsNone           PropertyConditionFlags = 0x0
	PropertyConditionFlagsIgnoreCase     PropertyConditionFlags = 0x1
	PropertyConditionFlagsMatchSubstring PropertyConditionFlags = 0x2
)

type TreeScope uint32

const (
	TreeScopeNone        TreeScope = 0x0
	TreeScopeElement     TreeScope = 0x1
	TreeScopeChildren    TreeScope = 0x2
	TreeScopeDescendants TreeScope = 0x4
	TreeScopeParent      TreeScope = 0x8
	TreeScopeAncestors   TreeScope = 0x10
	TreeScopeSubtree     TreeScope = TreeScopeElement | TreeScopeChildren | TreeScopeDescendants
)

type StructureChangeType uint32

const (
	StructureChangeTypeChildAdded StructureChangeType = iota
	StructureChangeTypeChildRemoved
	StructureChangeTypeChildrenInvalidated
	StructureChangeTypeChildrenBulkAdded
	StructureChangeTypeChildrenBulkRemoved
	StructureChangeTypeChildrenReordered
)

type WindowVisualState uint32

const (
	WindowVisualStateNormal WindowVisualState = iota
	WindowVisualStateMaximized
	WindowVisualStateMinimized
)

type WindowInteractionState uint32

const (
	WindowInteractionStateRunning WindowInteractionState = iota
	WindowInteractionStateClosing
	WindowInteractionStateReadyForUserInteraction
	WindowInteractionStateBlockedByModalWindow
	WindowInteractionStateNotResponding
)

type ScrollAmount uint32

const (
	ScrollAmountLargeDecrement ScrollAmount = iota
	ScrollAmountSmallDecrement
	ScrollAmountNoAmount
	ScrollAmountLargeIncrement
	ScrollAmountSmallIncrement
)

type ExpandCollapseState uint32

const (
	ExpandCollapseStateCollapsed ExpandCollapseState = iota
	ExpandCollapseStateExpanded
	ExpandCollapseStatePartiallyExpanded
	ExpandCollapseStateLeafNode
)

type DockPosition uint32

const (
	DockPositionTop DockPosition = iota
	DockPositionLeft
	DockPositionBottom
	DockPositionRight
	DockPositionFill
	DockPositionNone
)

type RowOrColumnMajor uint32

const (
	RowOrColumnMajorRowMajor RowOrColumnMajor = iota
	RowOrColumnMajorColumnMajor
	RowOrColumnMajorIndeterminate
)
