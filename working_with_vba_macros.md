---
layout: default
title: WORKING WITH VBA MACROS
---

### <a name="working_with_vba_macros" class="anchor"
    href="#working_with_vba_macros"><span class="octicon octicon-link"
    /></a>WORKING WITH VBA MACROS

An Excel `xlsm` file is exactly the same as a `xlsx` file except that
is includes an additional `vbaProject.bin` file which contains
functions and/or macros. Excel uses a different extension to
differentiate between the two file formats since files containing
macros are usually subject to additional security checks.

The `vbaProject.bin` file is a binary OLE COM container.  This was the
format used in older C<xls> versions of Excel prior to Excel 2007.
Unlike all of the other components of an xlsx/xlsm file the data isn't
stored in XML format.  Instead the functions and macros as stored as
pre-parsed binary format.  As such it wouldn't be feasible to define
macros and create a `vbaProject.bin` file from scratch (at least not
in the remaining lifespan and interest levels of the author).

Instead a workaround is used to extract `vbaProject.bin` files from
existing xlsm files and then add these to WriteXLSX files.


#### The extract_vba utility

The `extract_vba` utility is used to extract the `vbaProject.bin`
binary from an Excel 2007+ xlsm file.  The utility is included in the
WriteXLSX bin directory and is also installed as a standalone
executable file:

    $ extract_vba macro_file.xlsm
    Extracted: vbaProject.bin


#### Adding the VBA macros to a WriteXLSX file

Once the `vbaProject.bin` file has been extracted it can be added to
the WriteXLSX workbook using the `add_vba_project` method:

    workbook.add_vba_project('./vbaProject.bin')

If the VBA file contains functions you can then refer to them in
calculations using `write_formula`:

    worksheet.write_formula('A1', '=MyMortgageCalc(200000, 25)')

Excel files that contain functions and macros should use an `xlsm`
extension or else Excel will complain and possibly not open the file:

    workbook  = WriteXLSX.new('file.xlsm')

It is also possible to assign a macro to a button that is inserted
into a worksheet using the `insert_button` method:

    workbook  = WriteXLSX.new('file.xlsm')
    ...
    workbook.add_vba_project('./vbaProject.bin')

    worksheet.insert_button('C2', { :macro => 'my_macro' } )


It may be necessary to specify a more explicit macro name prefixed by
the workbook VBA name as follows:

    worksheet.insert_button('C2', { :macro => 'ThisWorkbook.my_macro' } )

See the [`macros.rb`](examples.html#macros) from the examples
directory for a working example.

Note: Button is the only VBA Control supported by WriteXLSX.  Due to
the large effort in implementation (1+ man months) it is unlikely that
any other form elements will be added in the future.


#### Setting the VBA codenames

VBA macros generally refer to workbook and worksheet objects.  If the
VBA codenames aren't specified then WriteXLSX will use the Excel
defaults of `ThisWorkbook` and `Sheet1`, `Sheet2` etc.

If the macro uses other codenames you can set them using the workbook
and worksheet `set_vba_name` methods as follows:

      workbook.set_vba_name('MyWorkbook')
      worksheet.set_vba_name('MySheet')

You can find the names that are used in the VBA editor or by unzipping
the `xlsm` file and grepping the files. The following shows how to do
that using libxml's xmllint
[http://xmlsoft.org/xmllint.html](http://xmlsoft.org/xmllint.html) to
format the XML for clarity:

    $ unzip myfile.xlsm -d myfile
    $ xmllint --format `find myfile -name "*.xml" | xargs` | grep "Pr.*codeName"

      <workbookPr codeName="MyWorkbook" defaultThemeVersion="124226"/>
      <sheetPr codeName="MySheet"/>


Note: This step is particularly important for macros created with
non-English versions of Excel.



#### What to do if it doesn't work

This feature should be considered experimental and there is no
guarantee that it will work in all cases. Some effort may be required
and some knowledge of VBA will certainly help.  If things don't work
out here are some things to try:

* Start with a simple macro file, ensure that it works and then add
  complexity.
* Try to extract the macros from an Excel 2007 file. The method should
  work with macros from later versions (it was also tested with Excel
  2010 macros). However there may be features in the macro files of
  more recent version of Excel that aren't backward compatible.
* Check the code names that macros use to refer to the workbook and
  worksheets (see the previous section above). In general VBA uses a
  code name of `ThisWorkbook` to refer to the current workbook and the
  sheet name (such as `Sheet1`) to refer to the worksheets. These are
  the defaults used by WriteXLSX. If the macro uses other names then
  you can specify these using the workbook and worksheet
  `set_vba_name` methods:

    workbook.set_vba_name('MyWorkbook')
    worksheet.set_vba_name('MySheet')
