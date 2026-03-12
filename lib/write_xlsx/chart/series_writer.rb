# -*- coding: utf-8 -*-
# frozen_string_literal: true

###############################################################################
#
# series_writer.rb - series and cache XML output helpers
#
###############################################################################

module Writexlsx
  class Chart
    module SeriesWriter
      private

      #
      # Write the series elements.
      #
      def write_series(series) # :nodoc:
        write_ser(series)
      end

      #
      # Write the <c:ser> element.
      #
      def write_ser(series) # :nodoc:
        @writer.tag_elements('c:ser') do
          write_ser_base(series) do
            write_c_invert_if_negative(series.invert_if_negative)
          end
          # Write the c:cat element.
          write_cat(series)
          # Write the c:val element.
          write_val(series)
          # Write the c:smooth element.
          write_c_smooth(series.smooth) if ptrue?(@smooth_allowed)
          # Write the c:extLst element.
          write_ext_lst_inverted_fill(series.inverted_color) if series.inverted_color
        end
        @series_index += 1
      end

      def write_ser_base(series)
        # Write the c:idx element.
        write_idx(@series_index)
        # Write the c:order element.
        write_order(@series_index)
        # Write the series name.
        write_series_name(series)
        # Write the c:spPr element.
        write_sp_pr(series)
        # Write the c:marker element.
        write_marker(series.marker)

        yield if block_given?

        # Write the c:dPt element.
        write_d_pt(series.points)
        # Write the c:dLbls element.
        write_d_lbls(series.labels)
        # Write the c:trendline element.
        write_trendline(series.trendline)
        # Write the c:errBars element.
        write_error_bars(series.error_bars)
      end

      #
      # Write the <c:idx> element.
      #
      def write_idx(val) # :nodoc:
        @writer.empty_tag('c:idx', [['val', val]])
      end

      #
      # Write the <c:order> element.
      #
      def write_order(val) # :nodoc:
        @writer.empty_tag('c:order', [['val', val]])
      end

      #
      # Write the series name.
      #
      def write_series_name(series) # :nodoc:
        if series.name_formula
          write_tx_formula(series.name_formula, series.name_id)
        elsif series.name
          write_tx_value(series.name)
        end
      end

      #
      # Write the <c:cat> element.
      #
      def write_cat(series) # :nodoc:
        formula = series.categories
        data_id = series.cat_data_id

        data = @formula_data[data_id] if data_id

        # Ignore <c:cat> elements for charts without category values.
        return unless formula

        @writer.tag_elements('c:cat') do
          # Check the type of cached data.
          type = get_data_type(data)
          if type == 'str'
            @cat_has_num_fmt = false
            # Write the c:strRef element.
            write_str_ref(formula, data, type)
          elsif type == 'multi_str'
            @cat_has_num_fmt = false
            # Write the c:multiLvLStrRef element.
            write_multi_lvl_str_ref(formula, data)
          else
            @cat_has_num_fmt = true
            # Write the c:numRef element.
            write_num_ref(formula, data, type)
          end
        end
      end

      #
      # Write the <c:val> element.
      #
      def write_val(series) # :nodoc:
        write_val_base(series.values, series.val_data_id, 'c:val')
      end

      def write_val_base(formula, data_id, tag) # :nodoc:
        data = @formula_data[data_id]

        @writer.tag_elements(tag) do
          # Unlike Cat axes data should only be numeric.

          # Write the c:numRef element.
          write_num_ref(formula, data, 'num')
        end
      end

      #
      # Write the <c:numRef> or <c:strRef> element.
      #
      def write_num_or_str_ref(tag, formula, data, type) # :nodoc:
        @writer.tag_elements(tag) do
          # Write the c:f element.
          write_series_formula(formula)
          if type == 'num'
            # Write the c:numCache element.
            write_num_cache(data)
          elsif type == 'str'
            # Write the c:strCache element.
            write_str_cache(data)
          end
        end
      end

      #
      # Write the <c:numRef> element.
      #
      def write_num_ref(formula, data, type) # :nodoc:
        write_num_or_str_ref('c:numRef', formula, data, type)
      end

      #
      # Write the <c:strRef> element.
      #
      def write_str_ref(formula, data, type) # :nodoc:
        write_num_or_str_ref('c:strRef', formula, data, type)
      end

      #
      # Write the <c:multiLvLStrRef> element.
      #
      def write_multi_lvl_str_ref(formula, data)
        return if data.empty?

        @writer.tag_elements('c:multiLvlStrRef') do
          # Write the c:f element.
          write_series_formula(formula)

          @writer.tag_elements('c:multiLvlStrCache') do
            # Write the c:ptCount element.
            write_pt_count(data.last.size)

            # Write the data arrays in reverse order.
            data.reverse.each do |arr|
              @writer.tag_elements('c:lvl') do
                # Write the c:pt element.
                arr.each_with_index { |a, i| write_pt(i, a) }
              end
            end
          end
        end
      end

      #
      # Write the <c:numLit> element for literal number list elements.
      #
      def write_num_lit(data)
        write_num_base('c:numLit', data)
      end

      #
      # Write the <c:f> element.
      #
      def write_series_formula(formula) # :nodoc:
        # Strip the leading '=' from the formula.
        formula = formula.sub(/^=/, '')

        @writer.data_element('c:f', formula)
      end

      #
      # Write the <c:numCache> element.
      #
      def write_num_cache(data) # :nodoc:
        write_num_base('c:numCache', data)
      end

      def write_num_base(tag, data)
        @writer.tag_elements(tag) do
          # Write the c:formatCode element.
          write_format_code('General')

          # Write the c:ptCount element.
          count = if data
                    data.size
                  else
                    0
                  end
          write_pt_count(count)

          data.each_with_index do |token, i|
            # Write non-numeric data as 0.
            if token &&
               token.to_s !~ /^([+-]?)(?=\d|\.\d)\d*(\.\d*)?([Ee]([+-]?\d+))?$/
              token = 0
            end

            # Write the c:pt element.
            write_pt(i, token)
          end
        end
      end

      #
      # Write the <c:strCache> element.
      #
      def write_str_cache(data) # :nodoc:
        @writer.tag_elements('c:strCache') do
          write_pt_count(data.size)
          write_pts(data)
        end
      end

      def write_pts(data)
        data.each_index { |i| write_pt(i, data[i]) }
      end

      #
      # Write the <c:formatCode> element.
      #
      def write_format_code(data) # :nodoc:
        @writer.data_element('c:formatCode', data)
      end

      #
      # Write the <c:ptCount> element.
      #
      def write_pt_count(val) # :nodoc:
        @writer.empty_tag('c:ptCount', [['val', val]])
      end

      #
      # Write the <c:pt> element.
      #
      def write_pt(idx, value) # :nodoc:
        return unless value

        attributes = [['idx', idx]]

        @writer.tag_elements('c:pt', attributes) { write_v(value) }
      end

      #
      # Write the <c:v> element.
      #
      def write_v(data) # :nodoc:
        @writer.data_element('c:v', data)
      end

      #
      # Write the <c:axId> elements for the primary or secondary axes.
      #
      def write_axis_ids(params)
        # Generate the axis ids.
        add_axis_ids(params)

        if params[:primary_axes] == 0
          # Write the axis ids for the secondary axes.
          write_axis_id(@axis2_ids[0])
          write_axis_id(@axis2_ids[1])
        else
          # Write the axis ids for the primary axes.
          write_axis_id(@axis_ids[0])
          write_axis_id(@axis_ids[1])
        end
      end

      #
      # Write the <c:axId> element.
      #
      def write_axis_id(val) # :nodoc:
        @writer.empty_tag('c:axId', [['val', val]])
      end
    end
  end
end
