# -*- coding: utf-8 -*-
# frozen_string_literal: true

###############################################################################
#
# DataWriting - A module for writing data to worksheet cells.
#
# Used in conjunction with WriteXLSX
#
# Copyright 2000-2011, John McNamara, jmcnamara@cpan.org
# Convert to ruby by Hideo NAKAMURA, nakamura.hideo@gmail.com
#

module Writexlsx
  class Worksheet
    module DataWriting
      include Writexlsx::Utility

      #
      # :call-seq:
      #  write(row, column [ , token [ , format ] ])
      #
      # Excel makes a distinction between data types such as strings, numbers,
      # blanks, formulas and hyperlinks. To simplify the process of writing
      # data the {#write()}[#method-i-write] method acts as a general alias for several more
      # specific methods:
      #
      def write(row, col, token = nil, format = nil, value1 = nil, value2 = nil)
        # Check for a cell reference in A1 notation and substitute row and column
        if (row_col_array = row_col_notation(row))
          _row, _col = row_col_array
          _token     = col
          _format    = token
          _value1    = format
          _value2    = value1
        else
          _row = row
          _col = col
          _token = token
          _format = format
          _value1 = value1
          _value2 = value2
        end
        _token ||= ''
        _token = _token.to_s if token.instance_of?(Time) || token.instance_of?(Date)

        if _format.respond_to?(:force_text_format?) && _format.force_text_format?
          write_string(_row, _col, _token, _format) # Force text format
        # Match an array ref.
        elsif _token.respond_to?(:to_ary)
          write_row(_row, _col, _token, _format, _value1, _value2)
        elsif _token.respond_to?(:coerce)  # Numeric
          write_number(_row, _col, _token, _format)
        elsif _token.respond_to?(:=~)  # String
          # Match integer with leading zero(s)
          if @leading_zeros && _token =~ /^0\d*$/
            write_string(_row, _col, _token, _format)
          elsif _token =~ /\A([+-]?)(?=\d|\.\d)\d*(\.\d*)?([Ee]([+-]?\d+))?\Z/
            write_number(_row, _col, _token, _format)
          # Match formula
          elsif _token =~ /^=/
            write_formula(_row, _col, _token, _format, _value1)
          # Match array formula
          elsif _token =~ /^\{=.*\}$/
            write_formula(_row, _col, _token, _format, _value1)
          # Match blank
          elsif _token == ''
            #        row_col_args.delete_at(2)     # remove the empty string from the parameter list
            write_blank(_row, _col, _format)
          elsif @workbook.strings_to_urls
            # https://, http://, ftp://, mailto:, internal:, external:
            url_token_re = %r{\A(?:(?:https?|ftp)://|mailto:|(?:in|ex)ternal:)}

            if _token =~ url_token_re
              write_url(_row, _col, _token, _format, _value1, _value2)
            else
              write_string(_row, _col, _token, _format)
            end
          else
            write_string(_row, _col, _token, _format)
          end
        else
          write_string(_row, _col, _token, _format)
        end
      end

      #
      # :call-seq:
      #   write_row(row, col, array [ , format ])
      #
      # Write a row of data starting from (row, col). Call write_col() if any of
      # the elements of the array are in turn array. This allows the writing
      # of 1D or 2D arrays of data in one go.
      #
      def write_row(row, col, tokens = nil, *options)
        # Check for a cell reference in A1 notation and substitute row and column
        if (row_col_array = row_col_notation(row))
          _row, _col = row_col_array
          _tokens    = col
          _options   = [tokens] + options
        else
          _row = row
          _col = col
          _tokens = tokens
          _options = options
        end
        raise "Not an array ref in call to write_row()$!" unless _tokens.respond_to?(:to_ary)

        _tokens.each do |_token|
          # Check for nested arrays
          if _token.respond_to?(:to_ary)
            write_col(_row, _col, _token, *_options)
          else
            write(_row, _col, _token, *_options)
          end
          _col += 1
        end
      end

      #
      # :call-seq:
      #   write_col(row, col, array [ , format ])
      #
      # Write a column of data starting from (row, col). Call write_row() if any of
      # the elements of the array are in turn array. This allows the writing
      # of 1D or 2D arrays of data in one go.
      #
      def write_col(row, col, tokens = nil, *options)
        if (row_col_array = row_col_notation(row))
          _row, _col = row_col_array
          _tokens    = col
          _options   = [tokens] + options if options
        else
          _row = row
          _col = col
          _tokens = tokens
          _options = options
        end

        _tokens.each do |_token|
          # write() will deal with any nested arrays
          write(_row, _col, _token, *_options)
          _row += 1
        end
      end

      #
      # :call-seq:
      #   write_comment(row, column, string, options = {})
      #
      # Write a comment to the specified row and column (zero indexed).
      #
      def write_comment(row, col, string = nil, options = nil)
        # Check for a cell reference in A1 notation and substitute row and column
        if (row_col_array = row_col_notation(row))
          _row, _col = row_col_array
          _string    = col
          _options   = string
        else
          _row = row
          _col = col
          _string = string
          _options = options
        end
        raise WriteXLSXInsufficientArgumentError if [_row, _col, _string].include?(nil)

        # Check that row and col are valid and store max and min values
        check_dimensions(_row, _col)
        store_row_col_max_min_values(_row, _col)

        @has_vml = true

        # Process the properties of the cell comment.
        @comments.add(@workbook, self, _row, _col, _string, _options)
      end

      #
      # :call-seq:
      #   write_number(row, column, number [ , format ])
      #
      # Write an integer or a float to the cell specified by row and column:
      #
      def write_number(row, col, number, format = nil)
        # Check for a cell reference in A1 notation and substitute row and column
        if (row_col_array = row_col_notation(row))
          _row, _col = row_col_array
          _number = col
          _format = number
        else
          _row = row
          _col = col
          _number = number
          _format = format
        end
        raise WriteXLSXInsufficientArgumentError if _row.nil? || _col.nil? || _number.nil?

        # Check that row and col are valid and store max and min values
        check_dimensions(_row, _col)
        store_row_col_max_min_values(_row, _col)

        store_data_to_table(NumberCellData.new(_number, _format), _row, _col)
      end

      #
      # :call-seq:
      #   write_string(row, column, string [, format ])
      #
      # Write a string to the specified row and column (zero indexed).
      # +format+ is optional.
      #
      def write_string(row, col, string = nil, format = nil)
        # Check for a cell reference in A1 notation and substitute row and column
        if (row_col_array = row_col_notation(row))
          _row, _col = row_col_array
          _string = col
          _format = string
        else
          _row = row
          _col = col
          _string = string
          _format = format
        end
        _string &&= _string.to_s
        raise WriteXLSXInsufficientArgumentError if _row.nil? || _col.nil? || _string.nil?

        # Check that row and col are valid and store max and min values
        check_dimensions(_row, _col)
        store_row_col_max_min_values(_row, _col)

        index = shared_string_index(_string.length > STR_MAX ? _string[0, STR_MAX] : _string)

        store_data_to_table(StringCellData.new(index, _format, _string), _row, _col)
      end

      #
      # :call-seq:
      #    write_rich_string(row, column, (string | format, string)+,  [,cell_format])
      #
      # The write_rich_string() method is used to write strings with multiple formats.
      # The method receives string fragments prefixed by format objects. The final
      # format object is used as the cell format.
      #
      def write_rich_string(row, col, *rich_strings)
        # Check for a cell reference in A1 notation and substitute row and column
        if (row_col_array = row_col_notation(row))
          _row, _col    = row_col_array
          _rich_strings = [col] + rich_strings
        else
          _row = row
          _col = col
          _rich_strings = rich_strings
        end
        raise WriteXLSXInsufficientArgumentError if [_row, _col, _rich_strings[0]].include?(nil)

        _xf = cell_format_of_rich_string(_rich_strings)

        # Check that row and col are valid and store max and min values
        check_dimensions(_row, _col)
        store_row_col_max_min_values(_row, _col)

        _fragments, _raw_string = rich_strings_fragments(_rich_strings)
        # can't allow 2 formats in a row
        return -4 unless _fragments

        # Check that the string si < 32767 chars.
        return 3 if _raw_string.size > @xls_strmax

        index = shared_string_index(xml_str_of_rich_string(_fragments))

        store_data_to_table(RichStringCellData.new(index, _xf, _raw_string), _row, _col)
      end

      #
      # :call-seq:
      #   write_blank(row, col, format)
      #
      # Write a blank cell to the specified row and column (zero indexed).
      # A blank cell is used to specify formatting without adding a string
      # or a number.
      #
      def write_blank(row, col, format = nil)
        # Check for a cell reference in A1 notation and substitute row and column
        if (row_col_array = row_col_notation(row))
          _row, _col = row_col_array
          _format = col
        else
          _row = row
          _col = col
          _format = format
        end
        raise WriteXLSXInsufficientArgumentError if [_row, _col].include?(nil)

        # Don't write a blank cell unless it has a format
        return unless _format

        # Check that row and col are valid and store max and min values
        check_dimensions(_row, _col)
        store_row_col_max_min_values(_row, _col)

        store_data_to_table(BlankCellData.new(_format), _row, _col)
      end

      #
      # Utility method to strip equal sign and array braces from a formula
      # and also expand out future and dynamic array formulas.
      #
      def prepare_formula(given_formula, expand_future_functions = nil)
        # Ignore empty/null formulas.
        return given_formula unless ptrue?(given_formula)

        # Remove array formula braces and the leading =.
        formula = given_formula.sub(/^\{(.*)\}$/, '\1').sub(/^=/, '')

        # # Don't expand formulas that the user has already expanded.
        return formula if formula =~ /_xlfn\./

        # Expand dynamic array formulas.
        formula = expand_formula(formula, 'ANCHORARRAY\(')
        formula = expand_formula(formula, 'BYCOL\(')
        formula = expand_formula(formula, 'BYROW\(')
        formula = expand_formula(formula, 'CHOOSECOLS\(')
        formula = expand_formula(formula, 'CHOOSEROWS\(')
        formula = expand_formula(formula, 'DROP\(')
        formula = expand_formula(formula, 'EXPAND\(')
        formula = expand_formula(formula, 'FILTER\(', '._xlws')
        formula = expand_formula(formula, 'HSTACK\(')
        formula = expand_formula(formula, 'LAMBDA\(')
        formula = expand_formula(formula, 'MAKEARRAY\(')
        formula = expand_formula(formula, 'MAP\(')
        formula = expand_formula(formula, 'RANDARRAY\(')
        formula = expand_formula(formula, 'REDUCE\(')
        formula = expand_formula(formula, 'SCAN\(')
        formula = expand_formula(formula, 'SEQUENCE\(')
        formula = expand_formula(formula, 'SINGLE\(')
        formula = expand_formula(formula, 'SORT\(', '._xlws')
        formula = expand_formula(formula, 'SORTBY\(')
        formula = expand_formula(formula, 'SWITCH\(')
        formula = expand_formula(formula, 'TAKE\(')
        formula = expand_formula(formula, 'TEXTSPLIT\(')
        formula = expand_formula(formula, 'TOCOL\(')
        formula = expand_formula(formula, 'TOROW\(')
        formula = expand_formula(formula, 'UNIQUE\(')
        formula = expand_formula(formula, 'VSTACK\(')
        formula = expand_formula(formula, 'WRAPCOLS\(')
        formula = expand_formula(formula, 'WRAPROWS\(')
        formula = expand_formula(formula, 'XLOOKUP\(')

        if !@use_future_functions && !ptrue?(expand_future_functions)
          return formula
        end

        # Future functions.
        formula = expand_formula(formula, 'ACOTH\(')
        formula = expand_formula(formula, 'ACOT\(')
        formula = expand_formula(formula, 'AGGREGATE\(')
        formula = expand_formula(formula, 'ARABIC\(')
        formula = expand_formula(formula, 'ARRAYTOTEXT\(')
        formula = expand_formula(formula, 'BASE\(')
        formula = expand_formula(formula, 'BETA.DIST\(')
        formula = expand_formula(formula, 'BETA.INV\(')
        formula = expand_formula(formula, 'BINOM.DIST.RANGE\(')
        formula = expand_formula(formula, 'BINOM.DIST\(')
        formula = expand_formula(formula, 'BINOM.INV\(')
        formula = expand_formula(formula, 'BITAND\(')
        formula = expand_formula(formula, 'BITLSHIFT\(')
        formula = expand_formula(formula, 'BITOR\(')
        formula = expand_formula(formula, 'BITRSHIFT\(')
        formula = expand_formula(formula, 'BITXOR\(')
        formula = expand_formula(formula, 'CEILING.MATH\(')
        formula = expand_formula(formula, 'CEILING.PRECISE\(')
        formula = expand_formula(formula, 'CHISQ.DIST.RT\(')
        formula = expand_formula(formula, 'CHISQ.DIST\(')
        formula = expand_formula(formula, 'CHISQ.INV.RT\(')
        formula = expand_formula(formula, 'CHISQ.INV\(')
        formula = expand_formula(formula, 'CHISQ.TEST\(')
        formula = expand_formula(formula, 'COMBINA\(')
        formula = expand_formula(formula, 'CONCAT\(')
        formula = expand_formula(formula, 'CONFIDENCE.NORM\(')
        formula = expand_formula(formula, 'CONFIDENCE.T\(')
        formula = expand_formula(formula, 'COTH\(')
        formula = expand_formula(formula, 'COT\(')
        formula = expand_formula(formula, 'COVARIANCE.P\(')
        formula = expand_formula(formula, 'COVARIANCE.S\(')
        formula = expand_formula(formula, 'CSCH\(')
        formula = expand_formula(formula, 'CSC\(')
        formula = expand_formula(formula, 'DAYS\(')
        formula = expand_formula(formula, 'DECIMAL\(')
        formula = expand_formula(formula, 'ERF.PRECISE\(')
        formula = expand_formula(formula, 'ERFC.PRECISE\(')
        formula = expand_formula(formula, 'EXPON.DIST\(')
        formula = expand_formula(formula, 'F.DIST.RT\(')
        formula = expand_formula(formula, 'F.DIST\(')
        formula = expand_formula(formula, 'F.INV.RT\(')
        formula = expand_formula(formula, 'F.INV\(')
        formula = expand_formula(formula, 'F.TEST\(')
        formula = expand_formula(formula, 'FILTERXML\(')
        formula = expand_formula(formula, 'FLOOR.MATH\(')
        formula = expand_formula(formula, 'FLOOR.PRECISE\(')
        formula = expand_formula(formula, 'FORECAST.ETS.CONFINT\(')
        formula = expand_formula(formula, 'FORECAST.ETS.SEASONALITY\(')
        formula = expand_formula(formula, 'FORECAST.ETS.STAT\(')
        formula = expand_formula(formula, 'FORECAST.ETS\(')
        formula = expand_formula(formula, 'FORECAST.LINEAR\(')
        formula = expand_formula(formula, 'FORMULATEXT\(')
        formula = expand_formula(formula, 'GAMMA.DIST\(')
        formula = expand_formula(formula, 'GAMMA.INV\(')
        formula = expand_formula(formula, 'GAMMALN.PRECISE\(')
        formula = expand_formula(formula, 'GAMMA\(')
        formula = expand_formula(formula, 'GAUSS\(')
        formula = expand_formula(formula, 'HYPGEOM.DIST\(')
        formula = expand_formula(formula, 'IFNA\(')
        formula = expand_formula(formula, 'IFS\(')
        formula = expand_formula(formula, 'IMAGE\(')
        formula = expand_formula(formula, 'IMCOSH\(')
        formula = expand_formula(formula, 'IMCOT\(')
        formula = expand_formula(formula, 'IMCSCH\(')
        formula = expand_formula(formula, 'IMCSC\(')
        formula = expand_formula(formula, 'IMSECH\(')
        formula = expand_formula(formula, 'IMSEC\(')
        formula = expand_formula(formula, 'IMSINH\(')
        formula = expand_formula(formula, 'IMTAN\(')
        formula = expand_formula(formula, 'ISFORMULA\(')
        formula = expand_formula(formula, 'ISOMITTED\(')
        formula = expand_formula(formula, 'ISOWEEKNUM\(')
        formula = expand_formula(formula, 'LET\(')
        formula = expand_formula(formula, 'LOGNORM.DIST\(')
        formula = expand_formula(formula, 'LOGNORM.INV\(')
        formula = expand_formula(formula, 'MAXIFS\(')
        formula = expand_formula(formula, 'MINIFS\(')
        formula = expand_formula(formula, 'MODE.MULT\(')
        formula = expand_formula(formula, 'MODE.SNGL\(')
        formula = expand_formula(formula, 'MUNIT\(')
        formula = expand_formula(formula, 'NEGBINOM.DIST\(')
        formula = expand_formula(formula, 'NORM.DIST\(')
        formula = expand_formula(formula, 'NORM.INV\(')
        formula = expand_formula(formula, 'NORM.S.DIST\(')
        formula = expand_formula(formula, 'NORM.S.INV\(')
        formula = expand_formula(formula, 'NUMBERVALUE\(')
        formula = expand_formula(formula, 'PDURATION\(')
        formula = expand_formula(formula, 'PERCENTILE.EXC\(')
        formula = expand_formula(formula, 'PERCENTILE.INC\(')
        formula = expand_formula(formula, 'PERCENTRANK.EXC\(')
        formula = expand_formula(formula, 'PERCENTRANK.INC\(')
        formula = expand_formula(formula, 'PERMUTATIONA\(')
        formula = expand_formula(formula, 'PHI\(')
        formula = expand_formula(formula, 'POISSON.DIST\(')
        formula = expand_formula(formula, 'QUARTILE.EXC\(')
        formula = expand_formula(formula, 'QUARTILE.INC\(')
        formula = expand_formula(formula, 'QUERYSTRING\(')
        formula = expand_formula(formula, 'RANK.AVG\(')
        formula = expand_formula(formula, 'RANK.EQ\(')
        formula = expand_formula(formula, 'RRI\(')
        formula = expand_formula(formula, 'SECH\(')
        formula = expand_formula(formula, 'SEC\(')
        formula = expand_formula(formula, 'SHEETS\(')
        formula = expand_formula(formula, 'SHEET\(')
        formula = expand_formula(formula, 'SKEW.P\(')
        formula = expand_formula(formula, 'STDEV.P\(')
        formula = expand_formula(formula, 'STDEV.S\(')
        formula = expand_formula(formula, 'T.DIST.2T\(')
        formula = expand_formula(formula, 'T.DIST.RT\(')
        formula = expand_formula(formula, 'T.DIST\(')
        formula = expand_formula(formula, 'T.INV.2T\(')
        formula = expand_formula(formula, 'T.INV\(')
        formula = expand_formula(formula, 'T.TEST\(')
        formula = expand_formula(formula, 'TEXTAFTER\(')
        formula = expand_formula(formula, 'TEXTBEFORE\(')
        formula = expand_formula(formula, 'TEXTJOIN\(')
        formula = expand_formula(formula, 'UNICHAR\(')
        formula = expand_formula(formula, 'UNICODE\(')
        formula = expand_formula(formula, 'VALUETOTEXT\(')
        formula = expand_formula(formula, 'VAR.P\(')
        formula = expand_formula(formula, 'VAR.S\(')
        formula = expand_formula(formula, 'WEBSERVICE\(')
        formula = expand_formula(formula, 'WEIBULL.DIST\(')
        formula = expand_formula(formula, 'XMATCH\(')
        formula = expand_formula(formula, 'XOR\(')
        expand_formula(formula, 'Z.TEST\(')
      end

      #
      # :call-seq:
      #   write_formula(row, column, formula [ , format [ , value ] ])
      #
      # Write a formula or function to the cell specified by +row+ and +column+:
      #
      def write_formula(row, col, formula = nil, format = nil, value = nil)
        # Check for a cell reference in A1 notation and substitute row and column
        if (row_col_array = row_col_notation(row))
          _row, _col = row_col_array
          _formula   = col
          _format    = formula
          _value     = format
        else
          _row = row
          _col = col
          _formula = formula
          _format = format
          _value = value
        end
        raise WriteXLSXInsufficientArgumentError if [_row, _col, _formula].include?(nil)

        # Check for dynamic array functions.
        regex = /\bANCHORARRAY\(|\bBYCOL\(|\bBYROW\(|\bCHOOSECOLS\(|\bCHOOSEROWS\(|\bDROP\(|\bEXPAND\(|\bFILTER\(|\bHSTACK\(|\bLAMBDA\(|\bMAKEARRAY\(|\bMAP\(|\bRANDARRAY\(|\bREDUCE\(|\bSCAN\(|\bSEQUENCE\(|\bSINGLE\(|\bSORT\(|\bSORTBY\(|\bSWITCH\(|\bTAKE\(|\bTEXTSPLIT\(|\bTOCOL\(|\bTOROW\(|\bUNIQUE\(|\bVSTACK\(|\bWRAPCOLS\(|\bWRAPROWS\(|\bXLOOKUP\(/
        if _formula =~ regex
          return write_dynamic_array_formula(
            _row, _col, _row, _col, _formula, _format, _value
          )
        end

        # Hand off array formulas.
        if _formula =~ /^\{=.*\}$/
          write_array_formula(_row, _col, _row, _col, _formula, _format, _value)
        else
          check_dimensions(_row, _col)
          store_row_col_max_min_values(_row, _col)
          _formula = prepare_formula(_formula)

          store_data_to_table(FormulaCellData.new(_formula, _format, _value), _row, _col)
        end
      end

      #
      # Internal method shared by the write_array_formula() and
      # write_dynamic_array_formula() methods.
      #
      def write_array_formula_base(type, *args)
        # Check for a cell reference in A1 notation and substitute row and column
        # Convert single cell to range
        if args.first.to_s =~ /^([A-Za-z]+[0-9]+)$/
          range = "#{::Regexp.last_match(1)}:#{::Regexp.last_match(1)}"
          params = [range] + args[1..-1]
        else
          params = args
        end

        if (row_col_array = row_col_notation(params.first))
          row1, col1, row2, col2 = row_col_array
          formula, xf, value = params[1..-1]
        else
          row1, col1, row2, col2, formula, xf, value = params
        end
        raise WriteXLSXInsufficientArgumentError if [row1, col1, row2, col2, formula].include?(nil)

        # Swap last row/col with first row/col as necessary
        row1, row2 = row2, row1 if row1 > row2
        col1, col2 = col2, col1 if col1 > col2

        # Check that row and col are valid and store max and min values
        check_dimensions(row1, col1)
        check_dimensions(row2, col2)
        store_row_col_max_min_values(row1, col1)
        store_row_col_max_min_values(row2, col2)

        # Define array range
        range = if row1 == row2 && col1 == col2
                  xl_rowcol_to_cell(row1, col1)
                else
                  "#{xl_rowcol_to_cell(row1, col1)}:#{xl_rowcol_to_cell(row2, col2)}"
                end

        # Modify the formula string, as needed.
        formula = prepare_formula(formula, 1)

        store_data_to_table(
          if type == 'a'
            FormulaArrayCellData.new(formula, xf, range, value)
          elsif type == 'd'
            DynamicFormulaArrayCellData.new(formula, xf, range, value)
          else
            raise "invalid type in write_array_formula_base()."
          end,
          row1, col1
        )

        # Pad out the rest of the area with formatted zeroes.
        (row1..row2).each do |row|
          (col1..col2).each do |col|
            next if row == row1 && col == col1

            write_number(row, col, 0, xf)
          end
        end
      end

      #
      # write_array_formula(row1, col1, row2, col2, formula, format)
      #
      # Write an array formula to the specified row and column (zero indexed).
      #
      def write_array_formula(row1, col1, row2 = nil, col2 = nil, formula = nil, format = nil, value = nil)
        write_array_formula_base('a', row1, col1, row2, col2, formula, format, value)
      end

      #
      # write_dynamic_array_formula(row1, col1, row2, col2, formula, format)
      #
      # Write a dynamic formula to the specified row and column (zero indexed).
      #
      def write_dynamic_array_formula(row1, col1, row2 = nil, col2 = nil, formula = nil, format = nil, value = nil)
        write_array_formula_base('d', row1, col1, row2, col2, formula, format, value)
        @has_dynamic_functions = true
      end

      #
      # write_boolean(row, col, val, format)
      #
      # Write a boolean value to the specified row and column (zero indexed).
      #
      def write_boolean(row, col, val = nil, format = nil)
        if (row_col_array = row_col_notation(row))
          _row, _col = row_col_array
          _val       = col
          _format    = val
        else
          _row = row
          _col = col
          _val = val
          _format = format
        end
        raise WriteXLSXInsufficientArgumentError if _row.nil? || _col.nil?

        _val = _val ? 1 : 0  # Boolean value.
        # xf : cell format.

        # Check that row and col are valid and store max and min values
        check_dimensions(_row, _col)
        store_row_col_max_min_values(_row, _col)

        store_data_to_table(BooleanCellData.new(_val, _format), _row, _col)
      end

      #
      # :call-seq:
      #   update_format_with_params(row, col, format_params)
      #
      # Update formatting of the cell to the specified row and column (zero indexed).
      #
      def update_format_with_params(row, col, params = nil)
        if (row_col_array = row_col_notation(row))
          _row, _col = row_col_array
          _params = args[1]
        else
          _row = row
          _col = col
          _params = params
        end
        raise WriteXLSXInsufficientArgumentError if _row.nil? || _col.nil? || _params.nil?

        # Check that row and col are valid and store max and min values
        check_dimensions(_row, _col)
        store_row_col_max_min_values(_row, _col)

        format = nil
        cell_data = nil
        if @cell_data_store[_row].nil? || @cell_data_store[_row][_col].nil?
          format = @workbook.add_format(_params)
          write_blank(_row, _col, format)
        else
          if @cell_data_store[_row][_col].xf.nil?
            format = @workbook.add_format(_params)
            cell_data = @cell_data_store[_row][_col]
          else
            format = @workbook.add_format
            cell_data = @cell_data_store[_row][_col]
            format.copy(cell_data.xf)
            format.set_format_properties(_params)
          end
          value = case cell_data
                  when FormulaCellData
                    "=#{cell_data.token}"
                  when FormulaArrayCellData
                    "{=#{cell_data.token}}"
                  when StringCellData
                    @workbook.shared_strings.string(cell_data.data[:sst_id])
                  else
                    cell_data.data
                  end
          write(_row, _col, value, format)
        end
      end

      #
      # :call-seq:
      #   write_url(row, column, url [ , format, label, tip ])
      #
      # Write a hyperlink to a URL in the cell specified by +row+ and +column+.
      # The hyperlink is comprised of two elements: the visible label and
      # the invisible link. The visible label is the same as the link unless
      # an alternative label is specified. The label parameter is optional.
      # The label is written using the {#write()}[#method-i-write] method. Therefore it is
      # possible to write strings, numbers or formulas as labels.
      #
      def write_url(row, col, url = nil, format = nil, str = nil, tip = nil, ignore_write_string = false)
        # Check for a cell reference in A1 notation and substitute row and column
        if (row_col_array = row_col_notation(row))
          _row, _col           = row_col_array
          _url                 = col
          _format              = url
          _str                 = format
          _tip                 = str
          _ignore_write_string = tip
        else
          _row                 = row
          _col                 = col
          _url                 = url
          _format              = format
          _str                 = str
          _tip                 = tip
          _ignore_write_string = ignore_write_string
        end

        _format, _str = _str, _format if _str.respond_to?(:xf_index) || (_format && !_format.respond_to?(:xf_index))
        raise WriteXLSXInsufficientArgumentError if [_row, _col, _url].include?(nil)

        # Check that row and col are valid and store max and min values
        check_dimensions(_row, _col)
        store_row_col_max_min_values(_row, _col)

        hyperlink = Hyperlink.factory(_url, _str, _tip, @max_url_length)
        store_hyperlink(_row, _col, hyperlink)

        raise "URL '#{url}' added but URL exceeds Excel's limit of 65,530 URLs per worksheet." if hyperlinks_count > 65_530

        # Add the default URL format.
        _format ||= @default_url_format

        # Write the hyperlink string.
        write_string(_row, _col, hyperlink.str, _format) unless _ignore_write_string
      end

      #
      # :call-seq:
      #   write_date_time (row, col, date_string [ , format ])
      #
      # Write a datetime string in ISO8601 "yyyy-mm-ddThh:mm:ss.ss" format as a
      # number representing an Excel date. format is optional.
      #
      def write_date_time(row, col, str, format = nil)
        # Check for a cell reference in A1 notation and substitute row and column
        if (row_col_array = row_col_notation(row))
          _row, _col = row_col_array
          _str       = col
          _format    = str
        else
          _row = row
          _col = col
          _str = str
          _format = format
        end
        raise WriteXLSXInsufficientArgumentError if [_row, _col, _str].include?(nil)

        # Check that row and col are valid and store max and min values
        check_dimensions(_row, _col)
        store_row_col_max_min_values(_row, _col)

        date_time = convert_date_time(_str)

        if date_time
          store_data_to_table(DateTimeCellData.new(date_time, _format), _row, _col)
        else
          # If the date isn't valid then write it as a string.
          write_string(_row, _col, _str, _format)
        end
      end

      #
      # Causes the write() method to treat integers with a leading zero as a string.
      # This ensures that any leading zeros such, as in zip codes, are maintained.
      #
      def keep_leading_zeros(flag = true)
        @leading_zeros = !!flag
      end

      #
      # merge_range(first_row, first_col, last_row, last_col, string, format)
      #
      # Merge a range of cells. The first cell should contain the data and the
      # others should be blank. All cells should contain the same format.
      #
      def merge_range(*args)
        if (row_col_array = row_col_notation(args.first))
          row_first, col_first, row_last, col_last = row_col_array
          string, format, *extra_args = args[1..-1]
        else
          row_first, col_first, row_last, col_last,
          string, format, *extra_args = args
        end

        raise "Incorrect number of arguments" if [row_first, col_first, row_last, col_last, format].include?(nil)
        raise "Fifth parameter must be a format object" unless format.respond_to?(:xf_index)
        raise "Can't merge single cell" if row_first == row_last && col_first == col_last

        # Swap last row/col with first row/col as necessary
        row_first,  row_last = row_last,  row_first  if row_first > row_last
        col_first, col_last = col_last, col_first if col_first > col_last

        # Check that the data range is valid and store the max and min values.
        check_dimensions(row_first, col_first)
        check_dimensions(row_last,  col_last)
        store_row_col_max_min_values(row_first, col_first)
        store_row_col_max_min_values(row_last,  col_last)

        # Store the merge range.
        @merge << [row_first, col_first, row_last, col_last]

        # Write the first cell
        write(row_first, col_first, string, format, *extra_args)

        # Pad out the rest of the area with formatted blank cells.
        write_formatted_blank_to_area(row_first, row_last, col_first, col_last, format)
      end

      #
      # Same as merge_range() above except the type of
      # {#write()}[#method-i-write] is specified.
      #
      def merge_range_type(type, *args)
        case type
        when 'array_formula', 'blank', 'rich_string'
          if (row_col_array = row_col_notation(args.first))
            row_first, col_first, row_last, col_last = row_col_array
            *others = args[1..-1]
          else
            row_first, col_first, row_last, col_last, *others = args
          end
          format = others.pop
        else
          if (row_col_array = row_col_notation(args.first))
            row_first, col_first, row_last, col_last = row_col_array
            token, format, *others = args[1..-1]
          else
            row_first, col_first, row_last, col_last,
            token, format, *others = args
          end
        end

        raise "Format object missing or in an incorrect position" unless format.respond_to?(:xf_index)
        raise "Can't merge single cell" if row_first == row_last && col_first == col_last

        # Swap last row/col with first row/col as necessary
        row_first, row_last = row_last, row_first if row_first > row_last
        col_first, col_last = col_last, col_first if col_first > col_last

        # Check that the data range is valid and store the max and min values.
        check_dimensions(row_first, col_first)
        check_dimensions(row_last,  col_last)
        store_row_col_max_min_values(row_first, col_first)
        store_row_col_max_min_values(row_last,  col_last)

        # Store the merge range.
        @merge << [row_first, col_first, row_last, col_last]

        # Write the first cell
        case type
        when 'blank', 'rich_string', 'array_formula'
          others << format
        end

        case type
        when 'string'
          write_string(row_first, col_first, token, format, *others)
        when 'number'
          write_number(row_first, col_first, token, format, *others)
        when 'blank'
          write_blank(row_first, col_first, *others)
        when 'date_time'
          write_date_time(row_first, col_first, token, format, *others)
        when 'rich_string'
          write_rich_string(row_first, col_first, *others)
        when 'url'
          write_url(row_first, col_first, token, format, *others)
        when 'formula'
          write_formula(row_first, col_first, token, format, *others)
        when 'array_formula'
          write_formula_array(row_first, col_first, *others)
        else
          raise "Unknown type '#{type}'"
        end

        # Pad out the rest of the area with formatted blank cells.
        write_formatted_blank_to_area(row_first, row_last, col_first, col_last, format)
      end

      #
      # :call-seq:
      #   repeat_formula(row, column, formula [ , format ])
      #
      # Deprecated. This is a writeexcel gem's method that is no longer
      # required by WriteXLSX.
      #
      def repeat_formula(row, col, formula, format, *pairs)
        # Check for a cell reference in A1 notation and substitute row and column.
        if (row_col_array = row_col_notation(row))
          _row, _col = row_col_array
          _formula   = col
          _format    = formula
          _pairs     = [format] + pairs
        else
          _row = row
          _col = col
          _formula = formula
          _format = format
          _pairs = pairs
        end
        raise WriteXLSXInsufficientArgumentError if [_row, _col].include?(nil)

        raise "Odd number of elements in pattern/replacement list" unless _pairs.size.even?
        raise "Not a valid formula" unless _formula.respond_to?(:to_ary)

        tokens  = _formula.join("\t").split("\t")
        raise "No tokens in formula" if tokens.empty?

        _value = nil
        if _pairs[-2] == 'result'
          _value = _pairs.pop
          _pairs.pop
        end
        until _pairs.empty?
          pattern = _pairs.shift
          replace = _pairs.shift

          tokens.each do |token|
            break if token.sub!(pattern, replace)
          end
        end
        _formula = tokens.join
        write_formula(_row, _col, _formula, _format, _value)
      end

      #
      # :call-seq:
      #   update_range_format_with_params(row_first, col_first, row_last, col_last, format_params)
      #
      # Update formatting of cells in range to the specified row and column (zero indexed).
      #
      def update_range_format_with_params(row_first, col_first, row_last = nil, col_last = nil, params = nil)
        if (row_col_array = row_col_notation(row_first))
          _row_first, _col_first, _row_last, _col_last = row_col_array
          params = args[1..-1]
        else
          _row_first = row_first
          _col_first = col_first
          _row_last  = row_last
          _col_last  = col_last
          _params    = params
        end

        raise WriteXLSXInsufficientArgumentError if [_row_first, _col_first, _row_last, _col_last, _params].include?(nil)

        # Swap last row/col with first row/col as necessary
        _row_first, _row_last = _row_last, _row_first if _row_first > _row_last
        _col_first, _col_last = _col_last, _col_first if _col_first > _col_last

        # Check that column number is valid and store the max value
        check_dimensions(_row_last, _col_last)
        store_row_col_max_min_values(_row_last, _col_last)

        (_row_first.._row_last).each do |row|
          (_col_first.._col_last).each do |col|
            update_format_with_params(row, col, _params)
          end
        end
      end

      private

      def store_hyperlink(row, col, hyperlink)
        @hyperlinks      ||= {}
        @hyperlinks[row] ||= {}
        @hyperlinks[row][col] = hyperlink
      end

      def hyperlinks_count
        @hyperlinks.keys.inject(0) { |s, n| s += @hyperlinks[n].keys.size }
      end

      # Pad out the rest of the area with formatted blank cells.
      def write_formatted_blank_to_area(row_first, row_last, col_first, col_last, format)
        (row_first..row_last).each do |row|
          (col_first..col_last).each do |col|
            next if row == row_first && col == col_first

            write_blank(row, col, format)
          end
        end
      end
    end
  end
end
