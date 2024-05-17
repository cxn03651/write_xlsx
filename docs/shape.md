---
layout: default
title: Shape
---
NAME

### <a name="shape" class="anchor" href="#shape"><span class="octicon octicon-link" /></a>Shape - A class for creating Excel Drawing shapes

#### <a name="synopsis" class="anchor" href="#synopsis"><span class="octicon octicon-link" /></a>SYNOPSIS

To create a simple Excel file containing shapes using WriteXLSX:

    require 'write_xlsx'

    workbook  = WriteXLSX.new( 'shape.xlsx' )
    worksheet = workbook.add_worksheet()

    # Add a default rectangle shape.
    rect = workbook.add_shape()

    # Add an ellipse with centered text.
    ellipse = workbook.add_shape(
        type: 'ellipse',
        text: "Hello\nWorld"
    )

    # Add a plus shape.
    plus = workbook.add_shape( type: 'plus')

    # Insert the shapes in the worksheet.
    worksheet.insert_shape( 'B3', rect )
    worksheet.insert_shape( 'C3', ellipse )
    worksheet.insert_shape( 'D3', plus )

### <a name="description" class="anchor" href="#description"><span class="octicon octicon-link" /></a>DESCRIPTION

A Shape object is created via the Workbook add_shape() method:

    shape_rect = workbook.add_shape( type: 'rect' )

Once the object is created it can be inserted into a worksheet using the
`insert_shape()` method:

    worksheet.insert_shape('A1', shape_rect)

A Shape can be inserted multiple times if required.

    worksheet.insert_shape('A1', shape_rect)
    worksheet.insert_shape('B2', shape_rect, 20, 30)

### <a name="methods" class="anchor" href="#methods"><span class="octicon octicon-link" /></a>METHODS

#### <a name="add_shape" class="anchor" href="#add_shape"><span class="octicon octicon-link" /></a>add_shape( properties )

The `add_shape()` Workbook method specifies the properties of the Shape in hash
`property: value` format:

    shape = workbook.add_shape( properties )

The available properties are shown below.

#### <a name="insert_shape" class="anchor" href="#insert_shape"><span class="octicon octicon-link" /></a>insert_shape( row, col, shape, x, y, scale_x, scale_y )

The `insert_shape()` Worksheet method sets the location and scale of the shape
object within the worksheet.

    # Insert the shape into the worksheet.
    worksheet.insert_shape( 'E2', shape )

Using the cell location and the `x` and `y` cell offsets it is possible to
position a shape anywhere on the canvas of a worksheet.

A more detailed explanation of the `insert_shape()` method is given
in the main WriteXLSX documentation.

#### <a name="shape_properties" class="anchor" href="#shape_properties"><span class="octicon octicon-link" /></a>SHAPE PROPERTIES

Any shape property can be queried or modified by the corresponding get/set method:

    ellipse = workbook.add_shape( properties )
    ellipse.set_type( 'plus' )    # No longer an ellipse!
    type = ellipse.get_type       # Find out what it really is.

Multiple shape properties may also be modified in one go by using the
`set_properties()` method:

    shape.set_properties( type: 'ellipse', text: 'Hello' )

The properties of a shape object that can be defined via `add_shape()`
are shown below.

##### <a name="shape_name" class="anchor" href="#shape_name"><span class="octicon octicon-link" /></a>:name

Defines the name of the shape.
This is an optional property and the shape will be given a default name
if not supplied.
The name is generally only used by Excel Macros to refer to the object.

##### <a name="shape_type" class="anchor" href="#shape_type"><span class="octicon octicon-link" /></a>:type

Defines the type of the object such as rect, ellipse or triangle:

    ellipse = workbook.add_shape( type: 'ellipse' )

The default type is rect.

The full list of available shapes is shown below.

See also the [`shapes_all.rb`](examples.html#shapes_all) program in the examples
directory of the distro.
It creates an example workbook with all supported shapes labelled with their
shape names.

    Basic Shapes
    -------------
    blockArc              can            chevron       cube          decagon
    diamond               dodecagon      donut         ellipse       funnel
    gear6                 gear9          heart         heptagon      hexagon
    homePlate             lightningBolt  line          lineInv       moon
    nonIsoscelesTrapezoid noSmoking      octagon       parallelogram pentagon
    pie                   pieWedge       plaque        rect          round1Rect
    round2DiagRect        round2SameRect roundRect     rtTriangle    smileyFace
    snip1Rect             snip2DiagRect  snip2SameRect snipRoundRect star10
    star12                star16         star24        star32        star4
    star5                 star6          star7         star8         sun
    teardrop              trapezoid      triangle

    Arrow Shapes
    -------------
    bentArrow        bentUpArrow       circularArrow     curvedDownArrow
    curvedLeftArrow  curvedRightArrow  curvedUpArrow     downArrow
    leftArrow        leftCircularArrow leftRightArrow    leftRightCircularArrow
    leftRightUpArrow leftUpArrow       notchedRightArrow quadArrow
    rightArrow       stripedRightArrow swooshArrow       upArrow
    upDownArrow      uturnArrow

    Connector Shapes
    ----------------
    bentConnector2   bentConnector3   bentConnector4
    bentConnector5   curvedConnector2 curvedConnector3
    curvedConnector4 curvedConnector5 straightConnector1

    Callout Shapes
    --------------
    accentBorderCallout1  accentBorderCallout2  accentBorderCallout3
    accentCallout1        accentCallout2        accentCallout3
    borderCallout1        borderCallout2        borderCallout3
    callout1              callout2              callout3
    cloudCallout          downArrowCallout      leftArrowCallout
    leftRightArrowCallout quadArrowCallout      rightArrowCallout
    upArrowCallout        upDownArrowCallout    wedgeEllipseCallout
    wedgeRectCallout      wedgeRoundRectCallout

    Flow Chart Shapes
    -----------------
    flowChartAlternateProcess  flowChartCollate        flowChartConnector
    flowChartDecision          flowChartDelay          flowChartDisplay
    flowChartDocument          flowChartExtract        flowChartInputOutput
    flowChartInternalStorage   flowChartMagneticDisk   flowChartMagneticDrum
    flowChartMagneticTape      flowChartManualInput    flowChartManualOperation
    flowChartMerge             flowChartMultidocument  flowChartOfflineStorage
    flowChartOffpageConnector  flowChartOnlineStorage  flowChartOr
    flowChartPredefinedProcess flowChartPreparation    flowChartProcess
    flowChartPunchedCard       flowChartPunchedTape    flowChartSort
    flowChartSummingJunction   flowChartTerminator

    Action Shapes
    -------------
    actionButtonBackPrevious actionButtonBeginning actionButtonBlank
    actionButtonDocument     actionButtonEnd       actionButtonForwardNext
    actionButtonHelp         actionButtonHome      actionButtonInformation
    actionButtonMovie        actionButtonReturn    actionButtonSound

    Chart Shapes
    ------------
    Not to be confused with Excel Charts.

    chartPlus chartStar chartX

    Math Shapes
    -----------
    mathDivide mathEqual mathMinus mathMultiply mathNotEqual mathPlus

    Stars and Banners
    -----------------
    arc            bevel          bracePair  bracketPair chord
    cloud          corner         diagStripe doubleWave  ellipseRibbon
    ellipseRibbon2 foldedCorner   frame      halfFrame   horizontalScroll
    irregularSeal1 irregularSeal2 leftBrace  leftBracket leftRightRibbon
    plus           ribbon         ribbon2    rightBrace  rightBracket
    verticalScroll wave

    Tab Shapes
    ----------
    cornerTabs plaqueTabs squareTabs

##### <a name="shape_text" class="anchor" href="#shape_text"><span class="octicon octicon-link" /></a>:text

This property is used to make the shape act like a text box.

    rect = workbook.add_shape( type: 'rect', text: "Hello\nWorld" )

The text is super-imposed over the shape.
The text can be wrapped using the newline character \n.

##### <a name="shape_id" class="anchor" href="#shape_id"><span class="octicon octicon-link" /></a>:id

Identification number for internal identification. This number will be
auto-assigned, if not assigned, or if it is a duplicate.

##### <a name="shape_format" class="anchor" href="#shape_format"><span class="octicon octicon-link" /></a>:format

Workbook format for decorating the shape text
(font family, size, and decoration).

##### <a name="start_start_index" class="anchor" href="#start_start_index"><span class="octicon octicon-link" /></a>:start, :start_index

Shape indices of the starting point for a connector and the index of the
connection. Index numbers are zero-based, start from the top dead centre
and are counted clockwise.

Indices are typically created for vertices and centre points of shapes.
They are the blue connection points that appear when connection shapes
are selected manually in Excel.

##### <a name="end_end_index" class="anchor" href="#end_end_index"><span class="octicon octicon-link" /></a>:end, :end_index

Same as above but for end points and end connections.

##### <a name="start_side_end_side" class="anchor" href="#start_side_end_side"><span class="octicon octicon-link" /></a>:start_side, :end_side

This is either the letter b or r for the bottom or right side of the shape
to be connected to and from.

If the start, start_index, and start_side parameters are defined for a
connection shape, the shape will be auto located and linked to the starting
and ending shapes respectively. This can be very useful for flow and
organisation charts.

##### <a name="flip_h_v" class="anchor" href="#flip_h_v"><span class="octicon octicon-link" /></a>:flip_h, :flip_v

Set this value to 1, to flip the shape horizontally and/or vertically.

##### <a name="rotarion" class="anchor" href="#rotation"><span class="octicon octicon-link" /></a>:rotation

Shape rotation, in degrees, from 0 to 360.

##### <a name="line_fill" class="anchor" href="#line_fill"><span class="octicon octicon-link" /></a>:line, :fill

Shape colour for the outline and fill. Colours may be specified as a colour
index, or in RGB format, i.e. AA00FF.

See [COLOURS IN EXCEL](colors.html#colors) in the main documentation for
more information.

##### <a name="line_type" class="anchor" href="#line_type"><span class="octicon octicon-link" /></a>:line_type

Line type for shape outline. The default is solid. The list of possible
values is:

    dash, sysDot, dashDot, lgDash, lgDashDot, lgDashDotDot, solid

##### <a name="valign_align" class="anchor" href="#valign_align"><span class="octicon octicon-link" /></a>:valign, :align

Text alignment within the shape.

Vertical alignment can be:

    Setting     Meaning
    =======     =======
    t           Top
    ctr         Centre
    b           Bottom

Horizontal alignment can be:

    Setting     Meaning
    =======     =======
    l           Left
    r           Right
    ctr         Centre
    just        Justified

The default is to centre both horizontally and vertically.

##### <a name="scale_x_y" class="anchor" href="#scale_x_y"><span class="octicon octicon-link" /></a>:scale_x, :scale_y

Scale factor in x and y dimension, for scaling the shape width and height.
The default value is 1.

Scaling may be set on the shape object or via `insert_shape()`.

##### <a name="adjustments" class="anchor" href="#adjustments"><span class="octicon octicon-link" /></a>:adjustments

Adjustment of shape vertices. Most shapes do not use this. For some shapes,
there is a single adjustment to modify the geometry.
For instance, the plus shape has one adjustment to control the width of the spokes.

Connectors can have a number of adjustments to control the shape routing.
Typically, a connector will have 3 to 5 handles for routing the shape.
The adjustment is in percent of the distance from the starting shape to the
ending shape, alternating between the x and y dimension.
Adjustments may be negative, to route the shape away from the endpoint.

##### <a name="stencil" class="anchor" href="#stencil"><span class="octicon octicon-link" /></a>:stencil

Shapes work in stencil mode by default. That is, once a shape is inserted,
its connection is separated from its master.
The master shape may be modified after an instance is inserted, and only
subsequent insertions will show the modifications.

This is helpful for Org charts, where an employee shape may be created once,
and then the text of the shape is modified for each employee.

The `insert_shape()` method returns a reference to the inserted shape (the child).

Stencil mode can be turned off, allowing for shape(s) to be modified after
insertion. In this case the `insert_shape()` method returns a reference to the
inserted shape (the master). This is not very useful for inserting multiple
shapes, since the x/y coordinates also gets modified.

### <a name="tips" class="anchor" href="#tips"><span class="octicon octicon-link" /></a>TIPS

Use worksheet.hide_gridlines(2) to prepare a blank canvas without gridlines.

Shapes do not need to fit on one page. Excel will split a large drawing into
multiple pages if required. Use the page break preview to show page
boundaries superimposed on the drawing.

Connected shapes will auto-locate in Excel if you move either the starting
shape or the ending shape separately. However, if you select both shapes
(lasso or control-click), the connector will move with it, and the shape
adjustments will not re-calculate.

### <a name="example" class="anchor" href="#example"><span class="octicon octicon-link" /></a>EXAMPLE

    require 'write_xlsx'

    workbook  = WriteXLSX.new( 'shape.xlsx' )
    worksheet = workbook.add_worksheet

    # Add a default rectangle shape.
    rect = workbook.add_shape

    # Add an ellipse with centered text.
    ellipse = workbook.add_shape(
        type: 'ellipse',
        text: "Hello\nWorld"
    )

    # Add a plus shape.
    plus = workbook.add_shape( type: 'plus')

    # Insert the shapes in the worksheet.
    worksheet.insert_shape( 'B3', rect )
    worksheet.insert_shape( 'C3', ellipse )
    worksheet.insert_shape( 'D3', plus )

See also the [`shapes_*.rb`](examples.html#shape1) program in the examples
directory of the distro.
