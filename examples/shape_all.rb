#######################################################################
#
# A simple example of how to use the WriteXLSX gem to
# add all shapes (as currently implemented) to an Excel xlsx file.
#
# The list at the end consists of all the shape types defined as
# ST_ShapeType in ECMA-376, Office Open XML File Formats Part 4.
#
# The grouping by worksheet name is for illustration only. It isn't
# part of the ECMA-376 standard.
#
# reverse(c), May 2012, John McNamara, jmcnamara@cpan.org
# converted to ruby by Hideo NAKAMURA, nakamura.hideo@gmail.com
#
require 'write_xlsx'

shapes_list = <<EOS
Action	actionButtonBackPrevious
Action	actionButtonBeginning
Action	actionButtonBlank
Action	actionButtonDocument
Action	actionButtonEnd
Action	actionButtonForwardNext
Action	actionButtonHelp
Action	actionButtonHome
Action	actionButtonInformation
Action	actionButtonMovie
Action	actionButtonReturn
Action	actionButtonSound
Arrow	bentArrow
Arrow	bentUpArrow
Arrow	circularArrow
Arrow	curvedDownArrow
Arrow	curvedLeftArrow
Arrow	curvedRightArrow
Arrow	curvedUpArrow
Arrow	downArrow
Arrow	leftArrow
Arrow	leftCircularArrow
Arrow	leftRightArrow
Arrow	leftRightCircularArrow
Arrow	leftRightUpArrow
Arrow	leftUpArrow
Arrow	notchedRightArrow
Arrow	quadArrow
Arrow	rightArrow
Arrow	stripedRightArrow
Arrow	swooshArrow
Arrow	upArrow
Arrow	upDownArrow
Arrow	uturnArrow
Basic	blockArc
Basic	can
Basic	chevron
Basic	cube
Basic	decagon
Basic	diamond
Basic	dodecagon
Basic	donut
Basic	ellipse
Basic	funnel
Basic	gear6
Basic	gear9
Basic	heart
Basic	heptagon
Basic	hexagon
Basic	homePlate
Basic	lightningBolt
Basic	line
Basic	lineInv
Basic	moon
Basic	nonIsoscelesTrapezoid
Basic	noSmoking
Basic	octagon
Basic	parallelogram
Basic	pentagon
Basic	pie
Basic	pieWedge
Basic	plaque
Basic	rect
Basic	round1Rect
Basic	round2DiagRect
Basic	round2SameRect
Basic	roundRect
Basic	rtTriangle
Basic	smileyFace
Basic	snip1Rect
Basic	snip2DiagRect
Basic	snip2SameRect
Basic	snipRoundRect
Basic	star10
Basic	star12
Basic	star16
Basic	star24
Basic	star32
Basic	star4
Basic	star5
Basic	star6
Basic	star7
Basic	star8
Basic	sun
Basic	teardrop
Basic	trapezoid
Basic	triangle
Callout	accentBorderCallout1
Callout	accentBorderCallout2
Callout	accentBorderCallout3
Callout	accentCallout1
Callout	accentCallout2
Callout	accentCallout3
Callout	borderCallout1
Callout	borderCallout2
Callout	borderCallout3
Callout	callout1
Callout	callout2
Callout	callout3
Callout	cloudCallout
Callout	downArrowCallout
Callout	leftArrowCallout
Callout	leftRightArrowCallout
Callout	quadArrowCallout
Callout	rightArrowCallout
Callout	upArrowCallout
Callout	upDownArrowCallout
Callout	wedgeEllipseCallout
Callout	wedgeRectCallout
Callout	wedgeRoundRectCallout
Chart	chartPlus
Chart	chartStar
Chart	chartX
Connector	bentConnector2
Connector	bentConnector3
Connector	bentConnector4
Connector	bentConnector5
Connector	curvedConnector2
Connector	curvedConnector3
Connector	curvedConnector4
Connector	curvedConnector5
Connector	straightConnector1
FlowChart	flowChartAlternateProcess
FlowChart	flowChartCollate
FlowChart	flowChartConnector
FlowChart	flowChartDecision
FlowChart	flowChartDelay
FlowChart	flowChartDisplay
FlowChart	flowChartDocument
FlowChart	flowChartExtract
FlowChart	flowChartInputOutput
FlowChart	flowChartInternalStorage
FlowChart	flowChartMagneticDisk
FlowChart	flowChartMagneticDrum
FlowChart	flowChartMagneticTape
FlowChart	flowChartManualInput
FlowChart	flowChartManualOperation
FlowChart	flowChartMerge
FlowChart	flowChartMultidocument
FlowChart	flowChartOfflineStorage
FlowChart	flowChartOffpageConnector
FlowChart	flowChartOnlineStorage
FlowChart	flowChartOr
FlowChart	flowChartPredefinedProcess
FlowChart	flowChartPreparation
FlowChart	flowChartProcess
FlowChart	flowChartPunchedCard
FlowChart	flowChartPunchedTape
FlowChart	flowChartSort
FlowChart	flowChartSummingJunction
FlowChart	flowChartTerminator
Math	mathDivide
Math	mathEqual
Math	mathMinus
Math	mathMultiply
Math	mathNotEqual
Math	mathPlus
Star_Banner	arc
Star_Banner	bevel
Star_Banner	bracePair
Star_Banner	bracketPair
Star_Banner	chord
Star_Banner	cloud
Star_Banner	corner
Star_Banner	diagStripe
Star_Banner	doubleWave
Star_Banner	ellipseRibbon
Star_Banner	ellipseRibbon2
Star_Banner	foldedCorner
Star_Banner	frame
Star_Banner	halfFrame
Star_Banner	horizontalScroll
Star_Banner	irregularSeal1
Star_Banner	irregularSeal2
Star_Banner	leftBrace
Star_Banner	leftBracket
Star_Banner	leftRightRibbon
Star_Banner	plus
Star_Banner	ribbon
Star_Banner	ribbon2
Star_Banner	rightBrace
Star_Banner	rightBracket
Star_Banner	verticalScroll
Star_Banner	wave
Tabs	cornerTabs
Tabs	plaqueTabs
Tabs	squareTabs
EOS

#
# main
#

workbook = WriteXLSX.new('shape_all.xlsx')

worksheet  = nil
last_sheet = ''
row        = 0

shapes_list.each_line do |line|
  line = line.chomp
  next unless line =~ /^\w/    # Skip blank lines and comments.

  sheet, name = line.split("\t")
  if last_sheet != sheet
    worksheet = workbook.add_worksheet(sheet)
    row       = 2
  end
  last_sheet = sheet
  shape      = workbook.add_shape(
    type:   name,
    text:   name,
    width:  90,
    height: 90
  )

  # Connectors can not have labels, so write the connector name in the cell
  # to the left.
  worksheet.write(row, 0, name) if sheet == 'Connector'
  worksheet.insert_shape(row, 2, shape, 0, 0)
  row += 5
end

workbook.close
