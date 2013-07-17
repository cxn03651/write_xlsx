---
layout: default
title: WriteXLSX
---
#### <a name="compatibility_with_writeexcel" class="anchor" href="#conpatibility_with_writeexcel"><span class="octicon octicon-link" /></a>Compatibility with WriteExcel

The WriteXLSX rubygem is a drop-in replacement for WriteExcel rubygem.

It supports all of the features of WriteExcel with some minor differences
noted below.

    Workbook Methods            Support
    ================            ======
    new()                       Yes
    add_worksheet()             Yes
    add_format()                Yes
    add_chart()                 Yes
    add_shape()                 Yes. Not in WriteExcel.
    add_vba_project()           Yes. Not in WriteExcel.
    close()                     Yes
    set_properties()            Yes
    define_name()               Yes
    set_tempdir()               Yes
    set_custom_color()          Yes
    sheets()                    Yes
    set_1904()                  Yes
    add_chart_ext()             Not supported. Not required in WriteXLSX.
    compatibility_mode()        Deprecated. Not required in WriteXLSX.
    set_codepage()              Deprecated. Not required in WriteXLSX.


    Worksheet Methods           Support
    =================           =======
    write()                     Yes
    write_number()              Yes
    write_string()              Yes
    write_rich_string()         Yes. Not in WriteExcel.
    write_blank()               Yes
    write_row()                 Yes
    write_col()                 Yes
    write_date_time()           Yes
    write_url()                 Yes
    write_formula()             Yes
    write_array_formula()       Yes. Not in WriteExcel.
    write_comment()             Yes
    show_comments()             Yes
    set_comments_author()       Yes
    insert_image()              Yes/Partial, see docs.
    insert_chart()              Yes
    insert_shape()              Yes. Not in WriteExcel.
    insert_button()             Yes. Not in WriteExcel.
    data_validation()           Yes
    conditional_formatting()    Yes. Not in WriteExcel.
    add_sparkline()             Yes. Not in WriteExcel.
    add_table()                 Yes. Not in WriteExcel.
    name()                      Yes
    activate()                  Yes
    select()                    Yes
    hide()                      Yes
    set_first_sheet()           Yes
    protect()                   Yes
    set_selection()             Yes
    set_row()                   Yes.
    set_column()                Yes.
    set_default_row()           Yes. Not in WriteExcel.
    outline_settings()          Yes
    freeze_panes()              Yes
    split_panes()               Yes
    merge_range()               Yes
    merge_range_type()          Yes. Not in WriteExcel.
    set_zoom()                  Yes
    right_to_left()             Yes
    hide_zero()                 Yes
    set_tab_color()             Yes
    autofilter()                Yes
    filter_column()             Yes
    filter_column_list()        Yes. Not in WriteExcel.
    write_utf16be_string()      Deprecated. Use utf8 strings instead.
    write_utf16le_string()      Deprecated. Use utf8 strings instead.
    store_formula()             Deprecated. See docs.
    repeat_formula()            Deprecated. See docs.
    write_url_range()           Not supported. Not required in WriteXLSX.

    Page Set-up Methods         Support
    ===================         =======
    set_landscape()             Yes
    set_portrait()              Yes
    set_page_view()             Yes
    set_paper()                 Yes
    center_horizontally()       Yes
    center_vertically()         Yes
    set_margins()               Yes
    set_header()                Yes
    set_footer()                Yes
    repeat_rows()               Yes
    repeat_columns()            Yes
    hide_gridlines()            Yes
    print_row_col_headers()     Yes
    print_area()                Yes
    print_across()              Yes
    fit_to_pages()              Yes
    set_start_page()            Yes
    set_print_scale()           Yes
    set_h_pagebreaks()          Yes
    set_v_pagebreaks()          Yes

    Format Methods              Support
    ==============              =======
    set_font()                  Yes
    set_size()                  Yes
    set_color()                 Yes
    set_bold()                  Yes
    set_italic()                Yes
    set_underline()             Yes
    set_font_strikeout()        Yes
    set_font_script()           Yes
    set_font_outline()          Yes
    set_font_shadow()           Yes
    set_num_format()            Yes
    set_locked()                Yes
    set_hidden()                Yes
    set_align()                 Yes
    set_rotation()              Yes
    set_text_wrap()             Yes
    set_text_justlast()         Yes
    set_center_across()         Yes
    set_indent()                Yes
    set_shrink()                Yes
    set_pattern()               Yes
    set_bg_color()              Yes
    set_fg_color()              Yes
    set_border()                Yes
    set_bottom()                Yes
    set_top()                   Yes
    set_left()                  Yes
    set_right()                 Yes
    set_border_color()          Yes
    set_bottom_color()          Yes
    set_top_color()             Yes
    set_left_color()            Yes
    set_right_color()           Yes
