---
layout: default
title: Workbook Method
---

### <a name="workbook" class="anchor" href="#workbook"><span class="octicon octicon-link" /></a>WORKBOOK METHODS

The WriteXLSX rubygem provides an object oriented interface to a new Excel workbook.
The following methods are available through a new workbook.

* [new](#new)
* [add_worksheet](#add_worksheet)
* [add_format](#add_format)
* [add_chart](#add_chart)
* [add_shape](#add_shape)
* [add_vba_project](#add_vba_project)
* [add_vba_name](#add_vba_name)
* [close](#close)
* [set_size](#set_size)
* [set_tab_ratio](#set_tab_ratio)
* [set_properties](#set_properties)
* [set_custom_properties](#set_custom_properties)
* [define_name](#define_name)
* [set_tempdir](#set_tempdir)
* [set_custom_color](#set_custom_color)
* [sheets](#sheets)
* [worksheet_by_name](#worksheet_by_name)
* [set_1904](#set_1904)
* [set_calc_mode](#set_calc_mode)
* [default_url_format](#default_url_format)

#### <a name="new" class="anchor" href="#new"><span class="octicon octicon-link" /></a>new

A new Excel workbook is created using the `new()` constructor which accepts either a filename or a IO object as a parameter.
The following example creates a new Excel file based on a filename:

    require 'write_xlsx'

    workbook  = WriteXLSX.new('filename.xlsx')
    worksheet = workbook.add_worksheet
    worksheet.write(0, 0, 'Hi Excel!')

Here are some other examples of using `new()` with filenames:

    workbook1 = WriteXLSX.new(filename)
    workbook2 = WriteXLSX.new('/tmp/filename.xlsx')
    workbook3 = WriteXLSX.new("c:\\tmp\\filename.xlsx")
    workbook4 = WriteXLSX.new('c:\tmp\filename.xlsx')

The last two examples demonstrates how to create a file on DOS or Windows
where it is necessary to either escape the directory separator \ or to use single quotes
to ensure that it isn't interpolated.

It is recommended that the filename uses the extension `.xlsx` rather than `.xls`
since the latter causes an Excel warning when used with the XLSX format.

You can pass a  IO object to the new constructor.:

    require 'write_xlsx'
    require 'stringio'

    io = StringIO.new
    workbook = WriteXLSX.new(io)
    # store sheets, data to the workbook
    .......
    workbook.close
    # You can get XLSX binary from io.string.

And you can also pass default format properties.

    workbook = WriteXLSX.new(filename, font: 'Courier New', size: 11)

See the [CELL FORMATTING][] section for more details about Format properties and how to set them.

You can also pass directory path in which write_xlsx store temporary files.

    workbook = WriteXLSX.new(filename, tempdir: './temp/', ...)

To write url/link string as string (not hyperlink), pass strings_as_urls option.

    workbook = WriteXLSX.new(filename, strings_as_urls: false, ...)

The `new()` constructor returns a Workbook object that you can use to add worksheets and store data.

#### <a name="add_worksheet" class="anchor" href="#add_worksheet"><span class="octicon octicon-link" /></a>add_worksheet(sheetname = '')

At least one worksheet should be added to a new workbook.
A worksheet is used to write data into cells:

    worksheet1 = workbook.add_worksheet               # Sheet1
    worksheet2 = workbook.add_worksheet('Foglio2')    # Foglio2
    worksheet3 = workbook.add_worksheet('Data')       # Data
    worksheet4 = workbook.add_worksheet               # Sheet4

If sheetname is not specified the default Excel convention will be followed, i.e. Sheet1, Sheet2, etc.

The worksheet name must be a valid Excel worksheet name, i.e:
* It must be less than 32 characters.
* It cannot contain any of the following characters: `[ ] : * ? / \`
* It cannot start or end with an apostrophe.
* It cannot be the same as an existing worksheet name (or a case insensitive variant).
* It cannot be the reserved name "History" (or a case insensitive variant).

See the [Excel worksheet naming rules](https://support.office.com/en-ie/article/rename-a-worksheet-3f1f7148-ee83-404d-8ef0-9ff99fbad1f9).

#### <a name="add_format" class="anchor" href="#add_format"><span class="octicon octicon-link" /></a>add_format(properties = {})

The `add_format()` method can be used to create new Format objects which are used to apply formatting to a cell.
You can either define the properties at creation time via a hash of property values or later via method calls.

    format1 = workbook.add_format(props_hash)   # Set properties at creation
    format2 = workbook.add_format               # Set properties later

See the [CELL FORMATTING][] section for more details about Format properties and how to set them.

#### <a name="add_chart" class="anchor" href="#add_chart"><span class="octicon octicon-link" /></a>add_chart(properties)

This method is use to create a new chart either as a standalone worksheet (the default)
or as an embeddable object that can be inserted into a worksheet via the
[insert_chart()][] Worksheet method.

    chart = workbook.add_chart(:type => 'column')

The properties that can be set are:

    :type     (required)
    :subtype  (optional)
    :name     (optional)
    :embedded (optional)

##### :type
This is a required parameter. It defines the type of chart that will be created.

    chart = workbook.add_chart(:type => 'line')

The available types are:

    area
    bar
    column
    line
    pie
    doughnut
    scatter
    stock

##### :subtype
Used to define a chart subtype where available.

    chart = workbook.add_chart(:type => 'bar', :subtype => 'stacked')

See the [Chart Documentation][] for a list of available chart subtypes.

##### :name

Set the name for the chart sheet.
The name property is optional and if it isn't supplied will default to Chart1 .. n.
The name must be a valid Excel worksheet name.
See [add_worksheet()](#add_worksheet) for more details on valid sheet names.
The name property can be omitted for embedded charts.

    chart = workbook.add_chart(:type => 'line', :name => 'Results Chart')

##### :embedded

Specifies that the Chart object will be inserted in a worksheet via the
[insert_chart()][] Worksheet method.
It is an error to try insert a Chart that doesn't have this flag set.

    chart = workbook.add_chart(:type => 'line', :embedded => 1)

    # Configure the chart.
    ...

    # Insert the chart into the a worksheet.
    worksheet.insert_chart('E2', chart)

See [Chart Documentation][] for details on how to configure the chart object once it is created.
See also the
[`chart_\*.rb`](examples.html#chart_area))
programs in the examples directory of the distro.

#### <a name="add_shape" class="anchor" href="#add_shape"><span class="octicon octicon-link" /></a>add_shape(properties)

The `add_shape()` method can be used to create new shapes that may be inserted into a worksheet.

You can either define the properties at creation time via a hash of property values or later via method calls.

    # Set properties at creation.
    plus = workbook.add_shape(
        :type   => 'plus',
        :id     => 3,
        :width  => pw,
        :height => ph
    )

    # Default rectangle shape. Set properties later.
    rect =  workbook.add_shape

See [Shape](shape.html#shape) for details on how to configure the shape object once it is created.

See also the
[`shape\*.rb`](examples.html#shape1)
programs in the examples directory of the distro.

#### <a name="add_vba_project" class="anchor" href="#add_vba_project"><span class="octicon octicon-link" /></a>add_vba_project( 'vbaProject.bin' )

The `add_vba_project()` method can be used to add macros or functions to an WriteXLSX file using a binary
VBA project file that has been extracted from an existing Excel xlsm file.

    workbook  = WriteXLSX.new('file.xlsm')

    workbook.add_vba_project('./vbaProject.bin')

The supplied extract_vba utility can be used to extract the required vbaProject.bin
file from an existing Excel file:

    $ extract_vba file.xlsm
    Extracted 'vbaProject.bin' successfully

Macros can be tied to buttons using the worksheet `insert_button()` method
(see the ["WORKSHEET METHODS"](worksheet.html#worksheet) section for details):

    worksheet.insert_button('C2', { :macro => 'my_macro' })

Note, Excel uses the file extension xlsm instead of xlsx for files that contain macros.
It is advisable to follow the same convention.

See also the
[`macros.rb`](examples.html#macros)
example file and the ["WORKING WITH VBA MACROS"](working_with_vba_macros.html#working_with_vba_macros).


#### <a name="add_vba_name" class="anchor" href="#add_vba_name"><span class="octicon octicon-link" /></a>add_vba_name

The `set_vba_name` method can be used to set the VBA codename for the workbook.
This is sometimes required when a `vbaProject macro` included via `add_vba_project` refers to the workbook.
The default Excel VBA name of `ThisWorkbook` is used if a user defined name isn't specified.
See also ["WORKING WITH VBA MACROS"](working_with_vba_macros.html#working_with_vba_macros).


#### <a name="close" class="anchor" href="#close"><span class="octicon octicon-link" /></a>close

The `close()` method can be used to explicitly close an Excel file.

    workbook.close

An explicit `close()` is required if the file must be closed prior to performing some external action
on it such as copying it, reading its size or attaching it to an email.

In general, if you create a file with a size of 0 bytes or you fail to create a file you need to call `close()`.

#### <a name="set_size" class="anchor" href="#set_size"><span class="octicon octicon-link" /></a>set_size

The `set_size` method can be used to set the size of a workbook window.

    workbook.set_size(1200, 800)

The Excel window size was used in Excel 2007 to define the width and height of
a workbook window within the Multiple Document Interface (MDI).
In later versions of Excel for Windows this interface was dropped.
This method is currectly only useful when setting the window size in Excel for Mac 2011.
The units are pixels and the default size is 1073 x 644.

Note, this doesn't equate exactly to the Excel for Mac pixel size
since it is based on the original Excel 2007 for Windows sizing.

#### <a name="set_tab_ratio" class="anchor" href="#set_tab_ratio"><span class="octicon octicon-link" /></a>set_tab_ratio

<!-- The `set_tab_ratio()` method ca -->n be used to set the ratio between worksheet tabs and the horizontal slider at the bottom of a workbook.
This can be increased to give more room to the tabs or reduced to increase the size of the horizontal slider:

    workbook.set_tab_ratio(75)

The default value in Excel is 60.

#### <a name="set_properties" class="anchor" href="#set_properties"><span class="octicon octicon-link" /></a>set_properties

The `set_properties` method can be used to set the document properties of the Excel file created
by WriteXLSX. These properties are visible when you use
the Office Button -> Prepare -> Properties option in Excel
and are also available to external applications that read or index windows files.

The properties should be passed in hash format as follows:

    workbook.set_properties(
        "title    => 'This is an example spreadsheet',
        "author   => 'John McNamara',
        "comments => 'Created with Ruby and WriteXLSX'
    )

The properties that can be set are:

    :title
    :subject
    :author
    :manager
    :company
    :category
    :keywords
    :comments
    :status
    :hyperlink_base

See also the
[`properties.rb`](examples.html#properties)
program in the examples directory of the distro.

#### <a name="set_custom_properties" class="anchor" href="#set_custom_properties"><span class="octicon octicon-link" /></a>set_custom_properties

The set_custom_property method can be used to set custom document properties not covered by the set_properties method above. These properties are visible when you use the < Office Button -> Prepare -> Properties -> Advanced Properties -> Custom > option in Excel and are also available to external applications that read or index Windows files.

The set_custom_property method takes 3 parameters:

    workbook.set_custom_property(name, value, type)

Where the available types are:

    text
    date
    number
    bool

For example:

    workbook.set_custom_property('Checked by',      'Eve',                  'text'  )
    workbook.set_custom_property('Date completed',  '2016-12-12T23:00:00Z', 'date'  )
    workbook.set_custom_property('Document number', '12345' ,               'number')
    workbook.set_custom_property('Reference',       '1.2345',               'number')
    workbook.set_custom_property('Has review',      1,                      'bool'  )
    workbook.set_custom_property('Has sign off',    0,                      'bool'  )
    workbook.set_custom_property('Department',      'some_string',          'text'  )
    workbook.set_custom_property('Scale',           '1.2345678901234',      'number')

Dates should by in ISO8601 <yyyy-mm-ddThh:mm:ss.sssZ> date format in Zulu time, as shown above.

The text and number types are optional since they can usually be inferred from the data:

    workbook.set_custom_property('Checked by', 'Eve'   )
    workbook.set_custom_property('Reference',  '1.2345')


The C<name> and C<value> (for text types) are limited to 255 characters by Excel.


#### <a name="define_name" class="anchor" href="#define_name"><span class="octicon octicon-link" /></a>define_name

This method is used to defined a name that can be used to represent a value,
a single cell or a range of cells in a workbook.

For example to set a global/workbook name:

    # Global/workbook names.
    workbook.define_name('Exchange_rate', '=0.96')
    workbook.define_name('Sales',         '=Sheet1!$G$1:$H$10')

It is also possible to define a local/worksheet name by prefixing the name
with the sheet name using the syntax sheetname!definedname:

    # Local/worksheet name.
    workbook.define_name('Sheet2!Sales',  '=Sheet2!$G$1:$G$10')

If the sheet name contains spaces or special characters you must enclose
it in single quotes like in Excel:

    workbook.define_name("'New Data'!Sales",  '=Sheet2!$G$1:$G$10')

See the
[`defined_name.rb`](examples.html#defined_name)
program in the examples dir of the distro.

Refer to the following to see Excel's syntax rules for defined names:
http://office.microsoft.com/en-001/excel-help/define-and-use-names-in-formulas-HA010147120.aspx#BMsyntax_rules_for_names

#### <a name="set_tempdir" class="anchor" href="#set_tempdir"><span class="octicon octicon-link" /></a>set_tempdir(tempdir)

WriteXLSX stores worksheet data in temporary files prior to assembling the final workbook.

If the default temporary file directory isn't accessible to your application,
or doesn't contain enough space,
you can specify an alternative location using the `set_tempdir()` method:

    workbook.set_tempdir('/tmp/writeexcel')
    workbook.set_tempdir('c:\windows\temp\writeexcel')

The directory for the temporary file must exist, `set_tempdir()` will not create a new directory.

#### <a name="set_custom_color" class="anchor" href="#set_custom_color"><span class="octicon octicon-link" /></a>set_custom_color(index, red, green, blue)

The method is maintained for backward compatibility with Writeexcel. WriteXLSX programs don't require this method and colours can be specified using a Html style `#RRGGBB` value, see [/WORKING WITH COLOURS][].

The WriteXLSX  will be extended at a later stage to support the newer, semi-infinite, palette.

#### <a name="sheets" class="anchor" href="#sheets"><span class="octicon octicon-link" /></a>sheets( 0, 1, ... )

The `sheets()` method returns a list, or a sliced list, of the worksheets in a workbook.

If no arguments are passed the method returns a list of all the worksheets in the workbook.
This is useful if you want to repeat an operation on each worksheet:

    workbook.sheets.each do |worksheet|
      print worksheet.get_name
    end

You can also specify a slice list to return one or more worksheet objects:

    worksheet = workbook.sheets(0)
    worksheet.write('A1', 'Hello')

Or since the return value from `sheets()` is a reference to a worksheet
object you can write the above example as:

    workbook.sheets(0).write('A1', 'Hello')

The following example returns the first and last worksheet in a workbook:

    workbook.sheets(0, -1).each do |worksheet|
        # Do something
    end

#### <a name="worksheet_by_name" class="anchor" href="#worksheet_by_name"><span class="octicon octicon-link" /></a>worksheet_by_name

The worksheet_by_name function return a worksheet or chartsheet object in the workbook using the sheetname:

    worksheet = workbook.worksheet_by_name('Sheet1')

#### <a name="set_1904" class="anchor" href="#set_1904"><span class="octicon octicon-link" /></a>set_1904()

Excel stores dates as real numbers where the integer part stores the number of days
since the epoch and the fractional part stores the percentage of the day.
The epoch can be either 1900 or 1904. Excel for Windows uses 1900 and Excel for Macintosh uses 1904.
However, Excel on either platform will convert automatically between one system and the other.

WriteXLSX stores dates in the 1900 format by default.
If you wish to change this you can call the `set_1904()` workbook method.
You can query the current value by calling the `get_1904()` workbook method.
This returns `false` for 1900 and `true` for 1904.

See also [DATES AND TIME IN EXCEL][] for more information about working with Excel's date system.

In general you probably won't need to use `set_1904()`.


#### <a name="set_calc_mode" class="anchor" href="#set_calc_mode"><span class="octicon octicon-link" /></a>set_calc_mode()

+Set the calculation mode for formulas in the workbook. This is mainly of use for workbooks with slow formulas where you want to allow the user to calculate them manually.

The mode parameter can be one of the following strings:

:auto

The default. Excel will re-calculate formulas when a formula or a value affecting the formula changes.

:manual

Only re-calculate formulas when the user requires it. Generally by pressing F9.

:auto_except_tables

Excel will automatically re-calculate formulas except for tables.


#### <a name="default_url_format" class="anchor" href="#default_url_format"><span class="octicon octicon-link" /></a>default_url_format

The `default_url_format()` method gets a copy of the default url format used when a user defined format isn't specified with the worksheet `write_url()` method. The format is the hyperlink style defined by Excel for the default theme:

    url_format = workbook.default_url_format



[CELL NOTATION]: worksheet.html#cell-notation
[CELL FORMATTING]: cell_formatting.html#cell_formatting
[COLOURS IN EXCEL]: colors.html#colors
[DATA VALIDATION IN EXCEL]: data_validation.html#data_validation
[DATES AND TIME IN EXCEL]: dates_and_time.html#dates_and_time
[Chart Documentation]: chart.html#chart
[FORMULAS AND FUNCTIONS IN EXCEL]: formulas_and_functions.html#formulas_and_functions
[CONDITIONAL FORMATTING IN EXCEL]: conditional_formatting.html#conditional_formatting
[SPARKLINES IN EXCEL]: sparklines.html#sparklines
[TABLES IN EXCEL]: tables.html#tables
[insert_chart()]: worksheet.html#insert_chart
