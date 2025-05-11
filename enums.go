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

type SupportedTextSelection uint32

const (
	SupportedTextSelectionNone SupportedTextSelection = iota
	SupportedTextSelectionSingle
	SupportedTextSelectionMultiple
)

type ToggleState uint32

const (
	ToggleStateOff ToggleState = iota
	ToggleStateOn
	ToggleStateIndeterminate
)

type SynchronizedInputType uint32

const (
	SynchronizedInputTypeKeyUp          SynchronizedInputType = 0x1
	SynchronizedInputTypeKeyDown        SynchronizedInputType = 0x2
	SynchronizedInputTypeLeftMouseUp    SynchronizedInputType = 0x4
	SynchronizedInputTypeLeftMouseDown  SynchronizedInputType = 0x8
	SynchronizedInputTypeRightMouseUp   SynchronizedInputType = 0x10
	SynchronizedInputTypeRightMouseDown SynchronizedInputType = 0x20
)

type ZoomUnit uint32

const (
	ZoomUnitNoAmount ZoomUnit = iota
	ZoomUnitLargeDecrement
	ZoomUnitSmallDecrement
	ZoomUnitLargeIncrement
	ZoomUnitSmallIncrement
)

type NavigateDirection uint32

const (
	NavigateDirectionParent NavigateDirection = iota
	NavigateDirectionNextSibling
	NavigateDirectionPreviousSibling
	NavigateDirectionFirstChild
	NavigateDirectionLastChild
)

type ProviderOptions uint32

const (
	ProviderOptionsClientSideProvider     ProviderOptions = 0x1
	ProviderOptionsServerSideProvider     ProviderOptions = 0x2
	ProviderOptionsNonClientAreaProvider  ProviderOptions = 0x4
	ProviderOptionsOverrideProvider       ProviderOptions = 0x8
	ProviderOptionsProviderOwnsSetFocus   ProviderOptions = 0x10
	ProviderOptionsUseComThreading        ProviderOptions = 0x20
	ProviderOptionsRefuseNonClientSupport ProviderOptions = 0x40
	ProviderOptionsHasNativeIAccessible   ProviderOptions = 0x80
	ProviderOptionsUseClientCoordinates   ProviderOptions = 0x100
)

type AutomationElementMode uint32

const (
	AutomationElementModeNone AutomationElementMode = iota
	AutomationElementModeFull
)

type TextPatternRangeEndpoint uint32

const (
	TextPatternRangeEndpointStart TextPatternRangeEndpoint = iota
	TextPatternRangeEndpointEnd
)

type TextUnit uint32

const (
	TextUnitCharacter TextUnit = iota
	TextUnitFormat
	TextUnitWord
	TextUnitLine
	TextUnitParagraph
	TextUnitPage
	TextUnitDocument
)
