# -*- coding: utf-8 -*-
# frozen_string_literal: true

###############################################################################
#
# axis_writer.rb - axis-specific XML output helpers
#
###############################################################################

module Writexlsx
  class Chart
    module AxisWriter
      private

      #
      # Write the <c:catAx> element. Usually the X axis.
      #
      def write_cat_axis(params) # :nodoc:
        x_axis   = params[:x_axis]
        y_axis   = params[:y_axis]
        axis_ids = params[:axis_ids]

        # if there are no axis_ids then we don't need to write this element
        return unless axis_ids
        return if axis_ids.empty?

        position  = @cat_axis_position
        is_y_axis = @horiz_cat_axis

        # Overwrite the default axis position with a user supplied value.
        position = x_axis.position || position

        @writer.tag_elements('c:catAx') do
          write_axis_id(axis_ids[0])
          # Write the c:scaling element.
          write_scaling(x_axis.reverse)

          write_delete(1) unless ptrue?(x_axis.visible)

          # Write the c:axPos element.
          write_axis_pos(position, y_axis.reverse)

          # Write the c:majorGridlines element.
          write_major_gridlines(x_axis.major_gridlines)

          # Write the c:minorGridlines element.
          write_minor_gridlines(x_axis.minor_gridlines)

          # Write the axis title elements.
          if x_axis.formula
            write_title_formula(x_axis, is_y_axis, @x_axis, x_axis.layout)
          elsif x_axis.name
            write_title_rich(x_axis, is_y_axis, x_axis.name_font, x_axis.layout)
          end

          # Write the c:numFmt element.
          write_cat_number_format(x_axis)

          # Write the c:majorTickMark element.
          write_major_tick_mark(x_axis.major_tick_mark)

          # Write the c:minorTickMark element.
          write_minor_tick_mark(x_axis.minor_tick_mark)

          # Write the c:tickLblPos element.
          write_tick_label_pos(x_axis.label_position)

          # Write the c:spPr element for the axis line.
          write_sp_pr(x_axis)

          # Write the axis font elements.
          write_axis_font(x_axis.num_font)

          # Write the c:crossAx element.
          write_cross_axis(axis_ids[1])

          write_crossing(y_axis.crossing) if @show_crosses || ptrue?(x_axis.visible)
          # Write the c:auto element.
          write_auto(1) unless x_axis.text_axis
          # Write the c:labelAlign element.
          write_label_align(x_axis.label_align)
          # Write the c:labelOffset element.
          write_label_offset(100)
          # Write the c:tickLblSkip element.
          write_tick_lbl_skip(x_axis.interval_unit)
          # Write the c:tickMarkSkip element.
          write_tick_mark_skip(x_axis.interval_tick)
        end
      end

      #
      # Write the <c:valAx> element. Usually the Y axis.
      #
      def write_val_axis(x_axis, y_axis, axis_ids, position = nil)
        return unless axis_ids && !axis_ids.empty?

        write_val_axis_base(
          x_axis, y_axis,
          axis_ids[0],
          axis_ids[1],
          y_axis.position || position || @val_axis_position
        )
      end
      public :write_val_axis

      def write_val_axis_base(x_axis, y_axis, axis_ids_0, axis_ids_1, position)  # :nodoc:
        @writer.tag_elements('c:valAx') do
          write_axis_id(axis_ids_1)

          # Write the c:scaling element.
          write_scaling_with_param(y_axis)

          write_delete(1) unless ptrue?(y_axis.visible)

          # Write the c:axPos element.
          write_axis_pos(position, x_axis.reverse)

          # Write the c:majorGridlines element.
          write_major_gridlines(y_axis.major_gridlines)

          # Write the c:minorGridlines element.
          write_minor_gridlines(y_axis.minor_gridlines)

          # Write the axis title elements.
          if y_axis.formula
            write_title_formula(y_axis, @horiz_val_axis, nil, y_axis.layout)
          elsif y_axis.name
            write_title_rich(y_axis, @horiz_val_axis, y_axis.name_font, y_axis.layout)
          end

          # Write the c:numberFormat element.
          write_number_format(y_axis)

          # Write the c:majorTickMark element.
          write_major_tick_mark(y_axis.major_tick_mark)

          # Write the c:minorTickMark element.
          write_minor_tick_mark(y_axis.minor_tick_mark)

          # Write the c:tickLblPos element.
          write_tick_label_pos(y_axis.label_position)

          # Write the c:spPr element for the axis line.
          write_sp_pr(y_axis)

          # Write the axis font elements.
          write_axis_font(y_axis.num_font)

          # Write the c:crossAx element.
          write_cross_axis(axis_ids_0)

          write_crossing(x_axis.crossing)

          # Write the c:crossBetween element.
          write_cross_between(x_axis.position_axis)

          # Write the c:majorUnit element.
          write_c_major_unit(y_axis.major_unit)

          # Write the c:minorUnit element.
          write_c_minor_unit(y_axis.minor_unit)

          # Write the c:dispUnits element.
          write_disp_units(y_axis.display_units, y_axis.display_units_visible)
        end
      end

      #
      # Write the <c:dateAx> element. Usually the X axis.
      #
      def write_date_axis(params)  # :nodoc:
        x_axis    = params[:x_axis]
        y_axis    = params[:y_axis]
        axis_ids  = params[:axis_ids]

        return unless axis_ids && !axis_ids.empty?

        position  = @cat_axis_position

        # Overwrite the default axis position with a user supplied value.
        position = x_axis.position || position

        @writer.tag_elements('c:dateAx') do
          write_axis_id(axis_ids[0])
          # Write the c:scaling element.
          write_scaling_with_param(x_axis)

          write_delete(1) unless ptrue?(x_axis.visible)

          # Write the c:axPos element.
          write_axis_pos(position, y_axis.reverse)

          # Write the c:majorGridlines element.
          write_major_gridlines(x_axis.major_gridlines)

          # Write the c:minorGridlines element.
          write_minor_gridlines(x_axis.minor_gridlines)

          # Write the axis title elements.
          if x_axis.formula
            write_title_formula(x_axis, nil, nil, x_axis.layout)
          elsif x_axis.name
            write_title_rich(x_axis, nil, x_axis.name_font, x_axis.layout)
          end
          # Write the c:numFmt element.
          write_number_format(x_axis)
          # Write the c:majorTickMark element.
          write_major_tick_mark(x_axis.major_tick_mark)

          # Write the c:tickLblPos element.
          write_tick_label_pos(x_axis.label_position)
          # Write the c:spPr element for the axis line.
          write_sp_pr(x_axis)
          # Write the font elements.
          write_axis_font(x_axis.num_font)
          # Write the c:crossAx element.
          write_cross_axis(axis_ids[1])

          write_crossing(y_axis.crossing) if @show_crosses || ptrue?(x_axis.visible)

          # Write the c:auto element.
          write_auto(1)
          # Write the c:labelOffset element.
          write_label_offset(100)
          # Write the c:tickLblSkip element.
          write_tick_lbl_skip(x_axis.interval_unit)
          # Write the c:tickMarkSkip element.
          write_tick_mark_skip(x_axis.interval_tick)
          # Write the c:majorUnit element.
          write_c_major_unit(x_axis.major_unit)
          # Write the c:majorTimeUnit element.
          write_c_major_time_unit(x_axis.major_unit_type) if x_axis.major_unit
          # Write the c:minorUnit element.
          write_c_minor_unit(x_axis.minor_unit)
          # Write the c:minorTimeUnit element.
          write_c_minor_time_unit(x_axis.minor_unit_type) if x_axis.minor_unit
        end
      end

      def write_crossing(crossing)
        # Note, the category crossing comes from the value axis.
        if [nil, 'max', 'min'].include?(crossing)
          # Write the c:crosses element.
          write_crosses(crossing)
        else
          # Write the c:crossesAt element.
          write_c_crosses_at(crossing)
        end
      end

      def write_scaling_with_param(param)
        write_scaling(
          param.reverse,
          param.min,
          param.max,
          param.log_base
        )
      end

      #
      # Write the <c:scaling> element.
      #
      def write_scaling(reverse, min = nil, max = nil, log_base = nil) # :nodoc:
        @writer.tag_elements('c:scaling') do
          # Write the c:logBase element.
          write_c_log_base(log_base)
          # Write the c:orientation element.
          write_orientation(reverse)
          # Write the c:max element.
          write_c_max(max)
          # Write the c:min element.
          write_c_min(min)
        end
      end

      #
      # Write the <c:logBase> element.
      #
      def write_c_log_base(val) # :nodoc:
        return unless ptrue?(val)

        @writer.empty_tag('c:logBase', [['val', val]])
      end

      #
      # Write the <c:orientation> element.
      #
      def write_orientation(reverse = nil) # :nodoc:
        val = ptrue?(reverse) ? 'maxMin' : 'minMax'

        @writer.empty_tag('c:orientation', [['val', val]])
      end

      #
      # Write the <c:max> element.
      #
      def write_c_max(max = nil) # :nodoc:
        @writer.empty_tag('c:max', [['val', max]]) if max
      end

      #
      # Write the <c:min> element.
      #
      def write_c_min(min = nil) # :nodoc:
        @writer.empty_tag('c:min', [['val', min]]) if min
      end

      #
      # Write the <c:axPos> element.
      #
      def write_axis_pos(val, reverse = false) # :nodoc:
        if reverse
          val = 'r' if val == 'l'
          val = 't' if val == 'b'
        end

        @writer.empty_tag('c:axPos', [['val', val]])
      end

      #
      # Write the <c:numberFormat> element. Note: It is assumed that if a user
      # defined number format is supplied (i.e., non-default) then the sourceLinked
      # attribute is 0. The user can override this if required.
      #

      def write_number_format(axis) # :nodoc:
        axis.write_number_format(@writer)
      end

      #
      # Write the <c:numFmt> element. Special case handler for category axes which
      # don't always have a number format.
      #
      def write_cat_number_format(axis)
        axis.write_cat_number_format(@writer, @cat_has_num_fmt)
      end

      #
      # Write the <c:majorTickMark> element.
      #
      def write_major_tick_mark(val)
        return unless ptrue?(val)

        @writer.empty_tag('c:majorTickMark', [['val', val]])
      end

      #
      # Write the <c:minorTickMark> element.
      #
      def write_minor_tick_mark(val)
        return unless ptrue?(val)

        @writer.empty_tag('c:minorTickMark', [['val', val]])
      end

      #
      # Write the <c:tickLblPos> element.
      #
      def write_tick_label_pos(val) # :nodoc:
        val ||= 'nextTo'
        val = 'nextTo' if val == 'next_to'

        @writer.empty_tag('c:tickLblPos', [['val', val]])
      end

      #
      # Write the <c:crossAx> element.
      #
      def write_cross_axis(val = 'autoZero') # :nodoc:
        @writer.empty_tag('c:crossAx', [['val', val]])
      end

      #
      # Write the <c:crosses> element.
      #
      def write_crosses(val) # :nodoc:
        val ||= 'autoZero'

        @writer.empty_tag('c:crosses', [['val', val]])
      end

      #
      # Write the <c:crossesAt> element.
      #
      def write_c_crosses_at(val) # :nodoc:
        @writer.empty_tag('c:crossesAt', [['val', val]])
      end

      #
      # Write the <c:auto> element.
      #
      def write_auto(val) # :nodoc:
        @writer.empty_tag('c:auto', [['val', val]])
      end

      #
      # Write the <c:labelAlign> element.
      #
      def write_label_align(val) # :nodoc:
        val ||= 'ctr'
        if val == 'right'
          val = 'r'
        elsif val == 'left'
          val = 'l'
        end
        @writer.empty_tag('c:lblAlgn', [['val', val]])
      end

      #
      # Write the <c:labelOffset> element.
      #
      def write_label_offset(val) # :nodoc:
        @writer.empty_tag('c:lblOffset', [['val', val]])
      end

      #
      # Write the <c:tickLblSkip> element.
      #
      def write_tick_lbl_skip(val) # :nodoc:
        return unless val

        @writer.empty_tag('c:tickLblSkip', [['val', val]])
      end

      #
      # Write the <c:tickMarkSkip> element.
      #
      def write_tick_mark_skip(val)  # :nodoc:
        return unless val

        @writer.empty_tag('c:tickMarkSkip', [['val', val]])
      end

      #
      # Write the <c:majorGridlines> element.
      #
      def write_major_gridlines(gridlines) # :nodoc:
        write_gridlines_base('c:majorGridlines', gridlines)
      end

      #
      # Write the <c:minorGridlines> element.
      #
      def write_minor_gridlines(gridlines)  # :nodoc:
        write_gridlines_base('c:minorGridlines', gridlines)
      end

      def write_gridlines_base(tag, gridlines)  # :nodoc:
        return unless gridlines
        return if gridlines.respond_to?(:[]) && !ptrue?(gridlines[:_visible])

        if gridlines.line_defined?
          @writer.tag_elements(tag) { write_sp_pr(gridlines) }
        else
          @writer.empty_tag(tag)
        end
      end

      #
      # Write the <c:crossBetween> element.
      #
      def write_cross_between(val = nil) # :nodoc:
        val ||= @cross_between

        @writer.empty_tag('c:crossBetween', [['val', val]])
      end

      #
      # Write the <c:majorUnit> element.
      #
      def write_c_major_unit(val = nil) # :nodoc:
        return unless val

        @writer.empty_tag('c:majorUnit', [['val', val]])
      end

      #
      # Write the <c:minorUnit> element.
      #
      def write_c_minor_unit(val = nil) # :nodoc:
        return unless val

        @writer.empty_tag('c:minorUnit', [['val', val]])
      end

      #
      # Write the <c:majorTimeUnit> element.
      #
      def write_c_major_time_unit(val) # :nodoc:
        val ||= 'days'

        @writer.empty_tag('c:majorTimeUnit', [['val', val]])
      end

      #
      # Write the <c:minorTimeUnit> element.
      #
      def write_c_minor_time_unit(val) # :nodoc:
        val ||= 'days'

        @writer.empty_tag('c:minorTimeUnit', [['val', val]])
      end

      #
      # Write the <c:dispUnits> element.
      #
      def write_disp_units(units, display)
        return unless ptrue?(units)

        attributes = [['val', units]]

        @writer.tag_elements('c:dispUnits') do
          @writer.empty_tag('c:builtInUnit', attributes)
          if ptrue?(display)
            @writer.tag_elements('c:dispUnitsLbl') do
              @writer.empty_tag('c:layout')
            end
          end
        end
      end

      #
      # Write the axis font elements.
      #
      def write_axis_font(font) # :nodoc:
        return unless font

        @writer.tag_elements('c:txPr') do
          write_a_body_pr(font[:_rotation])
          write_a_lst_style
          @writer.tag_elements('a:p') do
            write_a_p_pr_rich(font)
            write_a_end_para_rpr
          end
        end
      end
    end
  end
end
