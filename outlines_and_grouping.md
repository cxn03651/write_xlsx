---
layout: default
title: Outlines and Grouping
---
### <a name="outlines_and_grouping" class="anchor" href="#outlines_and_grouping"><span class="octicon octicon-link" /></a>OUTLINES AND GROUPING IN EXCEL

Excel allows you to group rows or columns so that they can be hidden or
displayed with a single mouse click.
This feature is referred to as outlines.

Outlines can reduce complex data down to a few salient sub-totals
or summaries.

This feature is best viewed in Excel but the following is an ASCII
representation of what a worksheet with three outlines might look like.
Rows 3-4 and rows 7-8 are grouped at level 2.
Rows 2-9 are grouped at level 1.
The lines at the left hand side are called outline level bars.

            ------------------------------------------
     1 2 3 |   |   A   |   B   |   C   |   D   |  ...
            ------------------------------------------
      _    | 1 |   A   |       |       |       |  ...
     |  _  | 2 |   B   |       |       |       |  ...
     | |   | 3 |  (C)  |       |       |       |  ...
     | |   | 4 |  (D)  |       |       |       |  ...
     | -   | 5 |   E   |       |       |       |  ...
     |  _  | 6 |   F   |       |       |       |  ...
     | |   | 7 |  (G)  |       |       |       |  ...
     | |   | 8 |  (H)  |       |       |       |  ...
     | -   | 9 |   I   |       |       |       |  ...
     -     | . |  ...  |  ...  |  ...  |  ...  |  ...

Clicking the minus sign on each of the level 2 outlines will collapse and hide
the data as shown in the next figure. The minus sign changes to a plus sign
to indicate that the data in the outline is hidden.

            ------------------------------------------
     1 2 3 |   |   A   |   B   |   C   |   D   |  ...
            ------------------------------------------
      _    | 1 |   A   |       |       |       |  ...
     |     | 2 |   B   |       |       |       |  ...
     | +   | 5 |   E   |       |       |       |  ...
     |     | 6 |   F   |       |       |       |  ...
     | +   | 9 |   I   |       |       |       |  ...
     -     | . |  ...  |  ...  |  ...  |  ...  |  ...

Clicking on the minus sign on the level 1 outline will collapse the remaining
rows as follows:

            ------------------------------------------
     1 2 3 |   |   A   |   B   |   C   |   D   |  ...
            ------------------------------------------
           | 1 |   A   |       |       |       |  ...
     +     | . |  ...  |  ...  |  ...  |  ...  |  ...

Grouping in WriteXLSX is achieved by setting the outline level via
the `set_row()` and `set_column()` worksheet methods:

    set_row(row, height, format, hidden, level, collapsed)
    set_column(first_col, last_col, width, format, hidden, level, collapsed)

The following example sets an outline level of 1 for rows 1 and 2
(zero-indexed) and columns B to G.
The parameters `height` and `XF` are assigned default values since they are
undefined:

    worksheet.set_row(1, nil, nil, 0, 1)
    worksheet.set_row(2, nil, nil, 0, 1)
    worksheet.set_column('B:G', nil, nil, 0, 1)

Excel allows up to 7 outline levels. Therefore the `level` parameter should be
in the range `0 <= level <= 7`.

Rows and columns can be collapsed by setting the `hidden` flag for the hidden
rows/columns and setting the `collapsed` flag for the row/column that has the
collapsed + symbol:

    worksheet.set_row(1, nil, nil, 1, 1)
    worksheet.set_row(2, nil, nil, 1, 1)
    worksheet.set_row(3, nil, nil, 0, 0, 1)          # Collapsed flag.

    worksheet.set_column('B:G', nil, nil, 1, 1)
    worksheet.set_column('H:H', nil, nil, 0, 0, 1)   # Collapsed flag.

Note: Setting the `collapsed` flag is particularly important for compatibility
with OpenOffice.org and Gnumeric.

For a more complete example see the
[`outline.rb`](examples.html#outline)
and
[`outline_collapsed.rb`](examples.html#outline_collapsed)
programs in the examples directory of the distro.

Some additional outline properties can be set via the
[`outline_settings()`](worksheet.html#outline_settings)
worksheet method, see above.
