---
layout: default
title: Conditional Formatting
---

### <a name="conditional_format" class="anchor" href="#conditional_format"><span class="octicon octicon-link" /></a>CONDITIONAL FORMATTING IN EXCEL

Conditional formatting is a feature of Excel which allows you to apply a format
to a cell or a range of cells based on a certain criteria.

For example the following criteria is used to highlight cells >= 50 in red in the
[`conditional_format.rb`](examples.html#conditional_format)
example from the distro:

    # Write a conditional format over a range.
    worksheet.conditional_formatting('B3:K12',
        {
            :type     => 'cell',
            :criteria => '>=',
            :value    => 50,
            :format   => format1
        }
    )

![The output from the above example](images/conditional_example.jpg)

#### <a name="conditional_formatting" class="anchor" href="#conditional_formatting"><span class="octicon octicon-link" /></a>conditional_formatting(row, col, { parameter => 'value', ... } )

The `conditional_formatting()` method is used to apply formatting based on user
defined criteria to an WriteXLSX file.

It can be applied to a single cell or a range of cells.
You can pass 3 parameters such as (row, col, \{...\}) or 5 parameters such as (first_row, first_col, last_row, last_col, \{...\}).
You can also use A1 style notation. For example:

    worksheet.conditional_formatting(0, 0,       {...})
    worksheet.conditional_formatting(0, 0, 4, 1, {...})

    # Which are the same as:

    worksheet.conditional_formatting('A1',       {...})
    worksheet.conditional_formatting('A1:B5',    {...})

See also the note about
["Cell notation"](worksheet.html#cell-notation)
for more information.

Using A1 style notation is is also possible to specify non-contiguous ranges,
separated by a comma. For example:

    worksheet.conditional_formatting('A1:D5,A8:D12', {...})

The last parameter in `conditional_formatting()` must be a hash ref containing
the parameters that describe the type and style of the data validation.
The main parameters are:

    :type
    :format
    :criteria
    :value
    :minimum
    :maximum

Other, less commonly used parameters are:

    :min_type
    :mid_type
    :max_type
    :min_value
    :mid_value
    :max_value
    :min_color
    :mid_color
    :max_color
    :bar_color

Additional parameters which are used for specific conditional format types
are shown in the relevant sections below.

##### <a name="type" class="anchor" href="#type"><span class="octicon octicon-link" /></a>:type

This parameter is passed in a hash ref to `conditional_formatting()`.

The `:type` parameter is used to set the type of conditional formatting that
you wish to apply. It is always required and it has no default value.
Allowable type values and their associated parameters are:

    Type            Parameters
    ====            ==========
    cell            criteria
                    value
                    minimum
                    maximum

    date            criteria
                    value
                    minimum
                    maximum

    time_period     criteria

    text            criteria
                    value

    average         criteria

    duplicate       (none)

    unique          (none)

    top             criteria
                    value

    bottom          criteria
                    value

    blanks          (none)

    no_blanks       (none)

    errors          (none)

    no_errors       (none)

    2_color_scale   (none)

    3_color_scale   (none)

    data_bar        (none)

    formula         criteria

All conditional formatting types have a format parameter, see below.
Other types and parameters such as icon sets will be added in time.

##### <a name="type_cell" class="anchor" href="#type_cell"><span class="octicon octicon-link" /></a>:type => 'cell'

This is the most common conditional formatting type.
It is used when a format is applied to a cell based on a simple criterion.
For example:

    worksheet.conditional_formatting('A1',
        {
            :type     => 'cell',
            :criteria => 'greater than',
            :value    => 5,
            :format   => $red_format,
        }
    )

Or, using the between criteria:

    worksheet.conditional_formatting('C1:C4',
        {
            :type     => 'cell',
            :criteria => 'between',
            :minimum  => 20,
            :maximum  => 30,
            :format   => $green_format,
        }
    )

##### <a name="criteria" class="anchor" href="#criteria"><span class="octicon octicon-link" /></a>:criteria

The `criteria` parameter is used to set the criteria by which the cell data
will be evaluated. It has no default value. The most common criteria as
applied to `{ :type => 'cell' }` are:

    'between'
    'not between'
    'equal to'                  |  '=='  |  '='
    'not equal to'              |  '!='  |  '<>'
    'greater than'              |  '>'
    'less than'                 |  '<'
    'greater than or equal to'  |  '>='
    'less than or equal to'     |  '<='

You can either use Excel's textual description strings,
in the first column above, or the more common symbolic alternatives.

Additional criteria which are specific to other conditional format types
are shown in the relevant sections below.

##### <a name="value" class="anchor" href="#"><span class="octicon octicon-link" /></a>:value

The value is generally used along with the criteria parameter to set the rule by which the cell data will be evaluated.

    :type     => 'cell',
    :criteria => '>',
    :value    => 5
    :format   => format,

The value property can also be an cell reference.

    :type     => 'cell',
    :criteria => '>',
    :value    => '$C$1',
    :format   => format,

##### <a name="format" class="anchor" href="#format"><span class="octicon octicon-link" /></a>:format

The `format` parameter is used to specify the format that will be applied to
the cell when the conditional formatting criterion is met.
The format is created using the `add_format()` method in the same way as cell
formats:

    format = workbook.add_format(:bold => 1, :italic => 1)

    worksheet.conditional_formatting('A1',
        {
            :type     => 'cell',
            :criteria => '>',
            :value    => 5
            ?format   => $format,
        }
    )

The conditional format follows the same rules as in Excel:
it is superimposed over the existing cell format and not all font and border
properties can be modified.
Font properties that can't be modified are font name,
font size, superscript and subscript.
The border property that cannot be modified is diagonal borders.

Excel specifies some default formats to be used with conditional formatting.
You can replicate them using the following WriteXLSX formats:

    # Light red fill with dark red text.

    format1 = workbook.add_format(
        :bg_color => '#FFC7CE',
        :color    => '#9C0006',
    )

    # Light yellow fill with dark yellow text.

    format2 = workbook.add_format(
        :bg_color => '#FFEB9C',
        :color    => '#9C6500',
    )

    # Green fill with dark green text.

    format3 = workbook.add_format(
        :bg_color => '#C6EFCE',
        :color    => '#006100',
    )

##### <a name="minimum" class="anchor" href="#minimum"><span class="octicon octicon-link" /></a>:minimum

The `minimum` parameter is used to set the lower limiting value when the
criteria is either 'between' or 'not between':

    :validate => 'integer',
    :criteria => 'between',
    :minimum  => 1,
    :maximum  => 100,

##### <a name="maximum" class="anchor" href="#maximum"><span class="octicon octicon-link" /></a>:maximum

The `maximum` parameter is used to set the upper limiting value when the
criteria is either 'between' or 'not between'. See the previous example.

##### <a name="type_date" class="anchor" href="#type_date"><span class="octicon octicon-link" /></a>:type => 'date'

The `date` type is the same as the `cell` type and uses the same criteria
and values. However it allows the `:value`, `:minimum` and `:maximum` properties to be
specified in the ISO8601
`yyyy-mm-ddThh:mm:ss.sss` date format which is detailed in the
[`write_date_time()`](worksheet.html#write_date_time) method.

    worksheet.conditional_formatting('A1:A4',
        {
            :type     => 'date',
            :criteria => 'greater than',
            :value    => '2011-01-01T',
            :format   => format,
        }
    )

##### <a name="type_time_period" class="anchor" href="#type_time_period"><span class="octicon octicon-link" /></a>:type => 'time_period'

The `time_period` type is used to specify Excel's "Dates Occurring" style
conditional format.

    worksheet.conditional_formatting('A1:A4',
        {
            :type     => 'time_period',
            :criteria => 'yesterday',
            :format   => $format,
        }
    )

The period is set in the `:criteria` and can have one of the following values:

        :criteria => 'yesterday',
        :criteria => 'today',
        :criteria => 'last 7 days',
        :criteria => 'last week',
        :criteria => 'this week',
        :criteria => 'next week',
        :criteria => 'last month',
        :criteria => 'this month',
        :criteria => 'next month'

##### <a name="type_text" class="anchor" href="#type_text"><span class="octicon octicon-link" /></a>:type => 'text'

The `text` type is used to specify Excel's "Specific Text" style conditional
format. It is used to do simple string matching using the `:criteria` and
`:value` parameters:

    worksheet.conditional_formatting('A1:A4',
        {
            :type     => 'text',
            :criteria => 'containing',
            :value    => 'foo',
            :format   => format,
        }
    )

The `:criteria` can have one of the following values:

    :criteria => 'containing',
    :criteria => 'not containing',
    :criteria => 'begins with',
    :criteria => 'ends with',

The `:value` parameter should be a string or single character.

##### <a name="type_average" class="anchor" href="#type_average"><span class="octicon octicon-link" /></a>:type => 'average'

The `average` type is used to specify Excel's "Average" style conditional format.

    worksheet.conditional_formatting('A1:A4',
        {
            ?type     => 'average',
            :criteria => 'above',
            ?format   => format,
        }
    )

The type of average for the conditional format range is specified by the `:criteria`:

    :criteria => 'above',
    :criteria => 'below',
    :criteria => 'equal or above',
    :criteria => 'equal or below',
    :criteria => '1 std dev above',
    :criteria => '1 std dev below',
    :criteria => '2 std dev above',
    :criteria => '2 std dev below',
    :criteria => '3 std dev above',
    :criteria => '3 std dev below',

##### <a name="type_duplicate" class="anchor" href="#type_duplicate"><span class="octicon octicon-link" /></a>type => 'duplicate'

The `duplicate` type is used to highlight duplicate cells in a range:

    worksheet.conditional_formatting('A1:A4',
        {
            :type     => 'duplicate',
            :format   => format,
        }
    )

##### <a name="type_unique" class="anchor" href="#type_unique"><span class="octicon octicon-link" /></a>:type => 'unique'

The `unique` type is used to highlight unique cells in a range:

    worksheet.conditional_formatting('A1:A4',
        {
            :type     => 'unique',
            :format   => format,
        }
    )

##### <a name="type_top" class="anchor" href="#type_top"><span class="octicon octicon-link" /></a>:type => 'top'

The `top` type is used to specify the top n values by number or percentage
in a range:

    worksheet.conditional_formatting('A1:A4',
        {
            :type     => 'top',
            :value    => 10,
            :format   => format,
        }
    )

The `:criteria` can be used to indicate that a percentage condition is required:

    worksheet.conditional_formatting('A1:A4',
        {
            :type     => 'top',
            :value    => 10,
            :criteria => '%',
            :format   => format,
        }
    )

##### <a name="type_bottom" class="anchor" href="#type_bottom"><span class="octicon octicon-link" /></a>:type => 'bottom'

The `bottom` type is used to specify the bottom n values by number or percentage
in a range.

It takes the same parameters as top, see above.

##### <a name="type_blanks" class="anchor" href="#type_blanks"><span class="octicon octicon-link" /></a>:type => 'blanks'

The `blanks` type is used to highlight blank cells in a range:

    worksheet.conditional_formatting('A1:A4',
        {
            :type     => 'blanks',
            :format   => format,
        }
    )

##### <a name="type_no_blanks" class="anchor" href="#type_no_blanks"><span class="octicon octicon-link" /></a>:type => 'no_blanks'

The `no_blanks` type is used to highlight non blank cells in a range:

    worksheet.conditional_formatting('A1:A4',
        {
            :type     => 'no_blanks',
            :format   => format,
        }
    )

##### <a name="type_errors" class="anchor" href="#type_errors"><span class="octicon octicon-link" /></a>type => 'errors'

The `errors` type is used to highlight error cells in a range:

    worksheet.conditional_formatting('A1:A4',
        {
            :type     => 'errors',
            :format   => format,
        }
    )

##### <a name="type_no_errors" class="anchor" href="#type_no_errors"><span class="octicon octicon-link" /></a>:type => 'no_errors'

The `no_errors` type is used to highlight non error cells in a range:

    worksheet.conditional_formatting('A1:A4',
        {
            :type     => 'no_errors',
            :format   => format,
        }
    )

##### <a name="type_2_color_scale" class="anchor" href="#type_2_color_scale"><span class="octicon octicon-link" /></a>:type => '2_color_scale'

The `2_color_scale` type is used to specify Excel's "2 Color Scale"
style conditional format.

    worksheet.conditional_formatting('A1:A12',
        {
            :type  => '2_color_scale',
        }
    )

This conditional type can be modified with `:min_type`, `:max_type`,
`:min_value`, `:max_value`, `:min_color` and `:max_color`, see below.

##### <a name="type_3_color_scale" class="anchor" href="#type_3_color_scale"><span class="octicon octicon-link" /></a>:type => '3_color_scale'

The `3_color_scale` type is used to specify Excel's "3 Color Scale"
style conditional format.

    worksheet.conditional_formatting('A1:A12',
        {
            :type  => '3_color_scale',
        }
    )

This conditional type can be modified with `:min_type`, `:mid_type`,
`:max_type`, `:min_value`, `:mid_value`, `:max_value`, `:min_color`,
`:mid_color` and `:max_color`, see below.

##### <a name="type_data_bar" class="anchor" href="#type_data_bar"><span class="octicon octicon-link" /></a>:type => 'data_bar'

The `data_bar` type is used to specify Excel's "Data Bar" style conditional format.

    worksheet.conditional_formatting('A1:A12',
        {
            :type  => 'data_bar',
        }
    )

This conditional type can be modified with `:min_type`, `:max_type`,
`:min_value`, `:max_value` and `:bar_color`, see below.

##### <a name="type_formula" class="anchor" href="#type_formula"><span class="octicon octicon-link" /></a>:type => 'formula'

The `formula` type is used to specify a conditional format based on a user
defined formula:

    worksheet.conditional_formatting('A1:A4',
        {
            :type     => 'formula',
            :criteria => '=$A$1 > 5',
            :format   => format,
        }
    )

The `formula` is specified in the criteria.

##### <a name="min_mid_max_type" class="anchor" href="#min_mid_max_type"><span class="octicon octicon-link" /></a>:min_type, :mid_type, :max_type

The `:min_type` and `:max_type` properties are available when the conditional
formatting type is `2_color_scale`, `3_color_scale` or `data_bar`.
The `:mid_type` is available for `3_color_scale`.
The properties are used as follows:

    worksheet.conditional_formatting('A1:A12',
        {
            :type      => '2_color_scale',
            :min_type  => 'percent',
            :max_type  => 'percent',
        }
    )

The available min/mid/max types are:

    num
    percent
    percentile
    formula

##### <a name="min_mid_max_value" class="anchor" href="#min_mid_max_value"><span class="octicon octicon-link" /></a>:min_value, :mid_value, :max_value

The `:min_value` and `:max_value` properties are available when the conditional
formatting type is `2_color_scale`, `3_color_scale` or `data_bar`.
The `:mid_value` is available for `3_color_scale`.
The properties are used as follows:

    worksheet.conditional_formatting('A1:A12',
        {
            :type       => '2_color_scale',
            :min_value  => 10,
            :max_value  => 90,
        }
    )

##### <a name="min_mid_max_bar_color" class="anchor" href="#min_mid_max_bar_color"><span class="octicon octicon-link" /></a>:min_color, :mid_color, :max_color, :bar_color

The `:min_color` and `:max_color` properties are available when the conditional
formatting type is `2_color_scale`, `3_color_scale` or `data_bar`.
The `:mid_color` is available for `3_color_scale`.
The properties are used as follows:

    worksheet.conditional_formatting('A1:A12',
        {
            :type      => '2_color_scale',
            :min_color => "#C5D9F1",
            :max_color => "#538ED5",
        }
    )

The color can be specifies as an WriteXLSX color index or, more usefully,
as a HTML style RGB hex number, as shown above.

#### <a name="conditional_formatting_examples" class="anchor" href="#conditional_formatting_examples"><span class="octicon octicon-link" /></a>Conditional Formatting Examples

Example 1. Highlight cells greater than an integer value.

    worksheet.conditional_formatting('A1:F10',
        {
            :type     => 'cell',
            :criteria => 'greater than',
            :value    => 5,
            :format   => format,
        }
    )

Example 2. Highlight cells greater than a value in a reference cell.

    worksheet.conditional_formatting('A1:F10',
        {
            :type     => 'cell',
            :criteria => 'greater than',
            :value    => '$H$1',
            :format   => format,
        }
    )

Example 3. Highlight cells greater than a certain date:

    worksheet.conditional_formatting('A1:F10',
        {
            :type     => 'date',
            :criteria => 'greater than',
            :value    => '2011-01-01T',
            :format   => format,
        }
    )

Example 4. Highlight cells with a date in the last seven days:

    worksheet.conditional_formatting('A1:F10',
        {
            :type     => 'time_period',
            :criteria => 'last 7 days',
            :format   => format,
        }
    )

Example 5. Highlight cells with strings starting with the letter b:

    worksheet.conditional_formatting('A1:F10',
        {
            :type     => 'text',
            :criteria => 'begins with',
            :value    => 'b',
            :format   => format,
        }
    )

Example 6. Highlight cells that are 1 std deviation above the average for the range:

    worksheet.conditional_formatting('A1:F10',
        {
            :type     => 'average',
            :format   => format,
        }
    )

Example 7. Highlight duplicate cells in a range:

    worksheet.conditional_formatting('A1:F10',
        {
            :type     => 'duplicate',
            :format   => format,
        }
    )

Example 8. Highlight unique cells in a range.

    worksheet.conditional_formatting('A1:F10',
        {
            :type     => 'unique',
            :format   => format,
        }
    )

Example 9. Highlight the top 10 cells.

    worksheet.conditional_formatting('A1:F10',
        {
            :type     => 'top',
            :value    => 10,
            :format   => format,
        }
    )

Example 10. Highlight blank cells.

    worksheet.conditional_formatting('A1:F10',
        {
            :type     => 'blanks',
            :format   => format,
        }
    )

See also the
[`conditional_format.rb`](examples.html#conditional_format)
example program in EXAMPLES.
