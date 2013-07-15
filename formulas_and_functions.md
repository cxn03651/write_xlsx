---
layout: default
title: Formulas and Functions
---
### <a name="formulas_and_functions" class="anchor" href="#formulas_and_functions"><span class="octicon octicon-link" /></a>FORMULAS AND FUNCTIONS IN EXCEL

##### <a name="introduction" class="anchor" href="#introduction"><span class="octicon octicon-link" /></a>Introduction

The following is a brief introduction to formulas and functions in Excel
and WriteXLSX.

A formula is a string that begins with an equals sign:

    '=A1+B1'
    '=AVERAGE(1, 2, 3)'

The formula can contain numbers, strings, boolean values, cell references,
cell ranges and functions. Named ranges are not supported.
Formulas should be written as they appear in Excel,
that is cells and functions must be in uppercase.

Cells in Excel are referenced using the A1 notation system where the column
is designated by a letter and the row by a number.
Columns range from A to XFD i.e. 0 to 16384, rows range from 1 to 1048576.

The Excel `$` notation in cell references is also supported.
This allows you to specify whether a row or column is relative or absolute.
This only has an effect if the cell is copied.
The following examples show relative and absolute values.

    '=A1'   # Column and row are relative
    '=$A1'  # Column is absolute and row is relative
    '=A$1'  # Column is relative and row is absolute
    '=$A$1' # Column and row are absolute

Formulas can also refer to cells in other worksheets of the current workbook.
For example:

    '=Sheet2!A1'
    '=Sheet2!A1:A5'
    '=Sheet2:Sheet3!A1'
    '=Sheet2:Sheet3!A1:A5'
    %q{='Test Data'!A1}
    %q{='Test Data1:Test Data2'!A1}

The sheet reference and the cell reference are separated by `!` the exclamation
mark symbol.
If worksheet names contain spaces, commas or parentheses then Excel requires
that the name is enclosed in single quotes as shown in the last two examples above.
In order to avoid using a lot of escape characters you can use the quote operator `%q{}`
to protect the quotes.
Only valid sheet names that have been added using the `add_worksheet()` method
can be used in formulas.
You cannot reference external workbooks.

The following table lists the operators that are available in Excel's formulas.
The majority of the operators are the same as Ruby's, differences are indicated:

    Arithmetic operators:
    =====================
    Operator  Meaning                   Example
       +      Addition                  1+2
       -      Subtraction               2-1
       *      Multiplication            2*3
       /      Division                  1/4
       ^      Exponentiation            2^3      # Equivalent to **
       -      Unary minus               -(1+2)
       %      Percent (Not modulus)     13%


    Comparison operators:
    =====================
    Operator  Meaning                   Example
        =     Equal to                  A1 =  B1 # Equivalent to ==
        <>    Not equal to              A1 <> B1 # Equivalent to !=
        >     Greater than              A1 >  B1
        <     Less than                 A1 <  B1
        >=    Greater than or equal to  A1 >= B1
        <=    Less than or equal to     A1 <= B1


    String operator:
    ================
    Operator  Meaning                   Example
        &     Concatenation             "Hello " & "World!" # [1]


    Reference operators:
    ====================
    Operator  Meaning                   Example
        :     Range operator            A1:A4               # [2]
        ,     Union operator            SUM(1, 2+2, B3)


    Notes:
    [1]: Equivalent to "Hello " + "World!" in Ruby.
    [2]: This range is equivalent to cells A1, A2, A3 and A4.

The range and comma operators can have different symbols in non-English
versions of Excel. These may be supported in a later version of
WriteXLSX. In the meantime European users of Excel take note:

    worksheet.write('A1', '=SUM(1; 2; 3)') # Wrong!!
    worksheet.write('A1', '=SUM(1, 2, 3)') # Okay

For a general introduction to Excel's formulas and an explanation
of the syntax of the function refer to the Excel help files or the following:
http://office.microsoft.com/en-us/assistance/CH062528031033.aspx.

If your formula doesn't work in WriteXLSX try the following:

1. Verify that the formula works in Excel.
2. Ensure that cell references and formula names are in uppercase.
3. Ensure that you are using ':' as the range operator, A1:A4.
4. Ensure that you are using ',' as the union operator, SUM(1,2,3).
5. If you verify that the formula works in Gnumeric, OpenOffice.org
   or LibreOffice, make sure to note items 2-4 above, since these
   applications are more flexible than Excel with formula syntax.
