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
versions of Excel, see below.

For a general introduction to Excel's formulas and an explanation
of the syntax of the function refer to the Excel help files or the following:
http://office.microsoft.com/en-us/assistance/CH062528031033.aspx.

In most cases a formula in Excel can be used directly in the write_formula method.
However, there are a few potential issues and differences that the user should be aware of.
These are explained in the following sections.

##### <a name="Non_US_Excel_functions_and_syntax" class="anchor" href="#Non_US_Excel_functions_and_syntax"><span class="octicon octicon-link" /></a>Non US Excel functions and syntax

 Excel stores formulas in the format of the US English version,
 regardless of the language or locale of the end-user's version of Excel.
 Therefore all formula function names written using WriteXLSX must be in English.

    worksheet.write_formula('A1', '=SUM(1, 2, 3)')    # OK
    worksheet.write_formula('A2', '=SOMME(1, 2, 3)')  # French, Error on load.

Also, formulas must be written with the US Style separator/range operator
which is a comma (not semi-colon).
Therefore a formula with multiple values should be written as follows:

    worksheet.write_formula('A1', '=SUM(1, 2, 3)')    # OK
    worksheet.write_formula('A2', '=SUM(1; 2; 3)')    # Semi-colon. Error on load.

If you have a non-English version of Excel you can use the following multi-lingual Formula Translator (https://en.excel-translator.de/language/) to help you convert the formula.
It can also replace semi-colon with commas.

##### <a name="Formulas_added_in_Excel_2010_and_later" class="anchor" href="#Formulas_added_in_Excel_2010_and_later"><span class="octicon octicon-link" /></a>Formulas added in Excel 2010 and later

Excel 2010 and later added functions which weren't defined in the original file specification.
These functions are reffered to by Microsoft as future functions.
Examples of these functions are ACOT, CHISQ.DIST.RT, CONFIDENCE.NORM, STDEV.P, STDEV.S and WORKDAY.INTL.

When written using write_formula() these functions need to be fully qualified with a _xlfn. (or other) prefix as they are shown the list below.
For example:

    worksheet.write_formula('A1', '=_xlfn.STDEV.S(B1:B10)')

They will appear without the prefix in Excel.

The following list is taken from MS XLSX extensions documentation on future functions: https://msdn.microsoft.com/en-us/library/dd907480%28v=office.12%29.aspx

    _xlfn.ACOT
    _xlfn.ACOTH
    _xlfn.AGGREGATE
    _xlfn.ARABIC
    _xlfn.BASE
    _xlfn.BETA.DIST
    _xlfn.BETA.INV
    _xlfn.BINOM.DIST
    _xlfn.BINOM.DIST.RANGE
    _xlfn.BINOM.INV
    _xlfn.BITAND
    _xlfn.BITLSHIFT
    _xlfn.BITOR
    _xlfn.BITRSHIFT
    _xlfn.BITXOR
    _xlfn.CEILING.MATH
    _xlfn.CEILING.PRECISE
    _xlfn.CHISQ.DIST
    _xlfn.CHISQ.DIST.RT
    _xlfn.CHISQ.INV
    _xlfn.CHISQ.INV.RT
    _xlfn.CHISQ.TEST
    _xlfn.COMBINA
    _xlfn.CONFIDENCE.NORM
    _xlfn.CONFIDENCE.T
    _xlfn.COT
    _xlfn.COTH
    _xlfn.COVARIANCE.P
    _xlfn.COVARIANCE.S
    _xlfn.CSC
    _xlfn.CSCH
    _xlfn.DAYS
    _xlfn.DECIMAL
    ECMA.CEILING
    _xlfn.ERF.PRECISE
    _xlfn.ERFC.PRECISE
    _xlfn.EXPON.DIST
    _xlfn.F.DIST
    _xlfn.F.DIST.RT
    _xlfn.F.INV
    _xlfn.F.INV.RT
    _xlfn.F.TEST
    _xlfn.FILTERXML
    _xlfn.FLOOR.MATH
    _xlfn.FLOOR.PRECISE
    _xlfn.FORECAST.ETS
    _xlfn.FORECAST.ETS.CONFINT
    _xlfn.FORECAST.ETS.SEASONALITY
    _xlfn.FORECAST.ETS.STAT
    _xlfn.FORECAST.LINEAR
    _xlfn.FORMULATEXT
    _xlfn.GAMMA
    _xlfn.GAMMA.DIST
    _xlfn.GAMMA.INV
    _xlfn.GAMMALN.PRECISE
    _xlfn.GAUSS
    _xlfn.HYPGEOM.DIST
    _xlfn.IFNA
    _xlfn.IMCOSH
    _xlfn.IMCOT
    _xlfn.IMCSC
    _xlfn.IMCSCH
    _xlfn.IMSEC
    _xlfn.IMSECH
    _xlfn.IMSINH
    _xlfn.IMTAN
    _xlfn.ISFORMULA
    ISO.CEILING
    _xlfn.ISOWEEKNUM
    _xlfn.LOGNORM.DIST
    _xlfn.LOGNORM.INV
    _xlfn.MODE.MULT
    _xlfn.MODE.SNGL
    _xlfn.MUNIT
    _xlfn.NEGBINOM.DIST
    NETWORKDAYS.INTL
    _xlfn.NORM.DIST
    _xlfn.NORM.INV
    _xlfn.NORM.S.DIST
    _xlfn.NORM.S.INV
    _xlfn.NUMBERVALUE
    _xlfn.PDURATION
    _xlfn.PERCENTILE.EXC
    _xlfn.PERCENTILE.INC
    _xlfn.PERCENTRANK.EXC
    _xlfn.PERCENTRANK.INC
    _xlfn.PERMUTATIONA
    _xlfn.PHI
    _xlfn.POISSON.DIST
    _xlfn.QUARTILE.EXC
    _xlfn.QUARTILE.INC
    _xlfn.QUERYSTRING
    _xlfn.RANK.AVG
    _xlfn.RANK.EQ
    _xlfn.RRI
    _xlfn.SEC
    _xlfn.SECH
    _xlfn.SHEET
    _xlfn.SHEETS
    _xlfn.SKEW.P
    _xlfn.STDEV.P
    _xlfn.STDEV.S
    _xlfn.T.DIST
    _xlfn.T.DIST.2T
    _xlfn.T.DIST.RT
    _xlfn.T.INV
    _xlfn.T.INV.2T
    _xlfn.T.TEST
    _xlfn.UNICHAR
    _xlfn.UNICODE
    _xlfn.VAR.P
    _xlfn.VAR.S
    _xlfn.WEBSERVICE
    _xlfn.WEIBULL.DIST
    WORKDAY.INTL
    _xlfn.XOR
    _xlfn.Z.TEST

##### <a name="Using_Tables_in_Formulas" class="anchor" href="#Using_Tables_in_Formulas"><span class="octicon octicon-link" /></a>Using Tables in Formulas

Worksheet tables can be aded with WriteXLSX using the add_table method.

    worksheet.add_table('B3:F7', {options})

By default tables are named Table1, Table2, etc., in the order that they are added.
However it can also be set by the user using the name parameter:

    worksheet.add_table('B3:F7', {'name': 'SalesData'})

If you need to know the name of the table, for example to use it in a formula,
you can get it as follows:

    table = worksheet.add_table('B3:F7')
    table_name = table.name

When used in a formula a table name such as TableX should be reffered to as TableX[]:

    worksheet.write_formula('A5', 'VLOOKUP("Sales", Table1[], 2, FALSE'))

##### <a name="Dealing_with_NAME_errors" class="anchor" href="#Dealing_with_NAME_errors"><span class="octicon octicon-link" /></a>Dealing with #NAME? errors

If there is an error in the syntax of a formula it is usually displayed in Excel as #NAME?.
If you encounter an error like this you can debug it as follows:

1. Ensure the formula is valild in Excel by copying and pasting it into a cell. Note, this should be done in Excel and not other applications such as OpenOffice or LibreOffice since they may have slightly difference syntax.
2. Ensure the formula is using comma separators instead of semi-colon, see [`"Non US Excel functions and syntax"`](#Non_US_Excel_functions_and_syntax) above.
3. Ensure the formula is in English, see [`"Non US Excel functions and syntax"`](#Non_US_Excel_functions_and_syntax) above.
4. Ensure that the formula doesn't contain an Excel 2010+ future function as listed in [`"Formulas added in Excel 2010 and later"`](#Formulas_added_in_Excel_2010_and_later) above. If it does then ensure that the correct prefix is used.

Finally if you have completed all the previous steps and still get a #NAME? error you can examine a valid Excel file to see what the correct syntax should be.
To do this you should create a valid formula in Excel and save the file.
You can then examine the XML in the unzipped file.

The following shows how to do that using Linux unzip and libxml's xmllint http://xmlsoft.org/xmllint.html to format the XML for clarity:

    $ unzipo myfile.xlsx -d myfile
    $ xmllint --format myfile/xl/worksheets/sheet1.xml | grep '<f>'

        <f>SUM(1, 2, 3)</f>

Formula Results

WriteXLSX doesn't calculate the result of a formula and instead stores the value 0 as the fomula result.
It then sets a global flag in the XLSX file to say that all formulas and functions should recalculated when the file is opened.

This is the method recommended in the Excel documentation and in general it works fine with spreadsheet applications.
However, applications that don't have a facility to calculate formulas will only display the 0 results.
Examples of such applications are Excel Viewer, PDF Converters, and some mobile device applications.

If required, it is also possible to specify the calculated result of the formula using the optional last `value` parameter in `write_formula`:

    worksheet.write_formula('A1', '=2+2', num_format, 4)

The `value` parameter can be a number, a string, a boolean string (`'TRUE'` or `'FALSE'`) or one of the following Excel error codes:

    #DIV/0!
    #N/A
    #NAME?
    #NULL!
    #NUM!
    #REF!
    #VALUE!

It is also possible to specify the calculated result of an array formula created with `write_array_formula`:

    # Specify the result for a single cell range.
    Worksheet.write_array_formula('A1:A1', '{=SUM(B1:C1*B2:C2)}', format, 2005)

However, using this parameter only writes a single value to the upper left cel in the result array.
For a multi-cell array formula where the results are required, the other result values can be specified by using `write_number` to write to the appropriate cell:

    #Specify the results for a multi cell range.
    worksheet.write_array_formula('A1:A3', '{=TREND(C1:C3,B1:B3)}', format, 15)
    worksheet.write_number('A2', 12, format)
    worksheet.write_number('A3', 14, format)
