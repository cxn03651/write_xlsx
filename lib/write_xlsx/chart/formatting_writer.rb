# -*- coding: utf-8 -*-
# frozen_string_literal: true

###############################################################################
#
# formatting_writer.rb - formatting, rich text, and label XML helpers
#
###############################################################################

module Writexlsx
  class Chart
    module FormattingWriter
      private

      #
      # Write the <c:title> element for a rich string.
      #
      def write_title_formula(title, is_y_axis = nil, axis = nil, layout = nil, overlay = nil) # :nodoc:
        @writer.tag_elements('c:title') do
          # Write the c:tx element.
          write_tx_formula(title.formula, axis ? axis.data_id : title.data_id)
          # Write the c:layout element.
          write_layout(layout, 'text')
          # Write the c:overlay element.
          write_overlay if overlay
          # Write the c:txPr element.
          write_tx_pr(axis ? axis.name_font : title.name_font, is_y_axis)
        end
      end

      #
      # Write the <c:tx> element.
      #
      def write_tx_rich(title, is_y_axis, font) # :nodoc:
        @writer.tag_elements('c:tx') do
          write_rich(title, font, is_y_axis)
        end
      end

      #
      # Write the <c:tx> element with a simple value such as for series names.
      #
      def write_tx_value(title) # :nodoc:
        @writer.tag_elements('c:tx') { write_v(title) }
      end

      #
      # Write the <c:tx> element.
      #
      def write_tx_formula(title, data_id) # :nodoc:
        data = @formula_data[data_id] if data_id

        @writer.tag_elements('c:tx') { write_str_ref(title, data, 'str') }
      end

      #
      # Write the <c:rich> element.
      #
      def write_rich(title, font, is_y_axis, ignore_rich_pr = false) # :nodoc:
        rotation = nil

        rotation = font[:_rotation] if font && font[:_rotation]
        @writer.tag_elements('c:rich') do
          # Write the a:bodyPr element.
          write_a_body_pr(rotation, is_y_axis)
          # Write the a:lstStyle element.
          write_a_lst_style
          # Write the a:p element.
          write_a_p_rich(title, font, ignore_rich_pr)
        end
      end

      #
      # Write the <a:p> element for rich string titles.
      #
      def write_a_p_rich(title, font, ignore_rich_pr) # :nodoc:
        @writer.tag_elements('a:p') do
          # Write the a:pPr element.
          write_a_p_pr_rich(font) unless ignore_rich_pr
          # Write the a:r element.
          write_a_r(title, font)
        end
      end

      #
      # Write the <a:pPr> element for rich string titles.
      #
      def write_a_p_pr_rich(font) # :nodoc:
        @writer.tag_elements('a:pPr') { write_a_def_rpr(font) }
      end

      #
      # Write the <a:r> element.
      #
      def write_a_r(title, font) # :nodoc:
        @writer.tag_elements('a:r') do
          # Write the a:rPr element.
          write_a_r_pr(font)
          # Write the a:t element.
          write_a_t(title.respond_to?(:name) ? title.name : title)
        end
      end

      #
      # Write the <a:rPr> element.
      #
      def write_a_r_pr(font) # :nodoc:
        attributes = [%w[lang en-US]]
        attr_font = get_font_style_attributes(font)
        attributes += attr_font unless attr_font.empty?

        write_def_rpr_r_pr_common(font, attributes, 'a:rPr')
      end

      #
      # Write the <a:t> element.
      #
      def write_a_t(title) # :nodoc:
        @writer.data_element('a:t', title)
      end

      #
      # Write the <c:marker> element.
      #
      def write_marker(marker = nil) # :nodoc:
        marker ||= @default_marker

        return unless ptrue?(marker)
        return if ptrue?(marker.automatic?)

        @writer.tag_elements('c:marker') do
          # Write the c:symbol element.
          write_symbol(marker.type)
          # Write the c:size element.
          size = marker.size
          write_marker_size(size) if ptrue?(size)
          # Write the c:spPr element.
          write_sp_pr(marker)
        end
      end

      #
      # Write the <c:marker> element without a sub-element.
      #
      def write_marker_value # :nodoc:
        return unless @default_marker

        @writer.empty_tag('c:marker', [['val', 1]])
      end

      #
      # Write the <c:size> element.
      #
      def write_marker_size(val) # :nodoc:
        @writer.empty_tag('c:size', [['val', val]])
      end

      #
      # Write the <c:symbol> element.
      #
      def write_symbol(val) # :nodoc:
        @writer.empty_tag('c:symbol', [['val', val]])
      end

      #
      # Write the <c:spPr> element.
      #
      def write_sp_pr(series) # :nodoc:
        return unless has_fill_formatting(series)

        line     = series_property(series, :line)
        fill     = series_property(series, :fill)
        pattern  = series_property(series, :pattern)
        gradient = series_property(series, :gradient)

        @writer.tag_elements('c:spPr') do
          # Write the fill elements for solid charts such as pie/doughnut and bar.
          if fill && fill[:_defined] != 0
            if ptrue?(fill[:none])
              # Write the a:noFill element.
              write_a_no_fill
            else
              # Write the a:solidFill element.
              write_a_solid_fill(fill)
            end
          end
          write_a_patt_fill(pattern) if ptrue?(pattern)
          if ptrue?(gradient)
            # Write the a:gradFill element.
            write_a_grad_fill(gradient)
          end
          # Write the a:ln element.
          write_a_ln(line) if line && ptrue?(line[:_defined])
        end
      end

      def series_property(object, property)
        if object.respond_to?(property)
          object.send(property)
        elsif object.respond_to?(:[])
          object[property]
        end
      end

      #
      # Write the <a:ln> element.
      #
      def write_a_ln(line) # :nodoc:
        attributes = []

        # Add the line width as an attribute.
        if line[:width]
          width = line[:width]
          # Round width to nearest 0.25, like Excel.
          width = ((width + 0.125) * 4).to_i / 4.0

          # Convert to internal units.
          width = (0.5 + (12700 * width)).to_i

          attributes << ['w', width]
        end

        if ptrue?(line[:none]) || ptrue?(line[:color]) || line[:dash_type]
          @writer.tag_elements('a:ln', attributes) do
            # Write the line fill.
            if ptrue?(line[:none])
              # Write the a:noFill element.
              write_a_no_fill
            elsif ptrue?(line[:color])
              # Write the a:solidFill element.
              write_a_solid_fill(line)
            end

            # Write the line/dash type.
            if line[:dash_type]
              # Write the a:prstDash element.
              write_a_prst_dash(line[:dash_type])
            end
          end
        else
          @writer.empty_tag('a:ln', attributes)
        end
      end

      #
      # Write the <a:noFill> element.
      #
      def write_a_no_fill # :nodoc:
        @writer.empty_tag('a:noFill')
      end

      #
      # Write the <a:alpha> element.
      #
      def write_a_alpha(val)
        val = (100 - val.to_i) * 1000

        @writer.empty_tag('a:alpha', [['val', val]])
      end

      #
      # Write the <a:prstDash> element.
      #
      def write_a_prst_dash(val) # :nodoc:
        @writer.empty_tag('a:prstDash', [['val', val]])
      end

      #
      # Write the <a:gradFill> element.
      #
      def write_a_grad_fill(gradient)
        attributes = [
          %w[flip none],
          ['rotWithShape', 1]
        ]
        attributes = [] if gradient[:type] == 'linear'

        @writer.tag_elements('a:gradFill', attributes) do
          # Write the a:gsLst element.
          write_a_gs_lst(gradient)

          if gradient[:type] == 'linear'
            # Write the a:lin element.
            write_a_lin(gradient[:angle])
          else
            # Write the a:path element.
            write_a_path(gradient[:type])

            # Write the a:tileRect element.
            write_a_tile_rect(gradient[:type])
          end
        end
      end

      #
      # Write the <a:gsLst> element.
      #
      def write_a_gs_lst(gradient)
        positions = gradient[:positions]
        colors    = gradient[:colors]

        @writer.tag_elements('a:gsLst') do
          (0..(colors.size - 1)).each do |i|
            pos = (positions[i] * 1000).to_i

            attributes = [['pos', pos]]
            @writer.tag_elements('a:gs', attributes) do
              color = color(colors[i])

              # Write the a:srgbClr element.
              # TODO: Wait for a feature request to support transparency.
              write_a_srgb_clr(color)
            end
          end
        end
      end

      #
      # Write the <a:lin> element.
      #
      def write_a_lin(angle)
        scaled = 0

        angle = (60000 * angle).to_i

        attributes = [
          ['ang',    angle],
          ['scaled', scaled]
        ]

        @writer.empty_tag('a:lin', attributes)
      end

      #
      # Write the <a:path> element.
      #
      def write_a_path(type)
        attributes = [['path', type]]

        @writer.tag_elements('a:path', attributes) do
          # Write the a:fillToRect element.
          write_a_fill_to_rect(type)
        end
      end

      #
      # Write the <a:fillToRect> element.
      #
      def write_a_fill_to_rect(type)
        attributes = if type == 'shape'
                       [
                         ['l', 50000],
                         ['t', 50000],
                         ['r', 50000],
                         ['b', 50000]
                       ]
                     else
                       [
                         ['l', 100000],
                         ['t', 100000]
                       ]
                     end

        @writer.empty_tag('a:fillToRect', attributes)
      end

      #
      # Write the <a:tileRect> element.
      #
      def write_a_tile_rect(type)
        attributes = if type == 'shape'
                       []
                     else
                       [
                         ['r', -100000],
                         ['b', -100000]
                       ]
                     end

        @writer.empty_tag('a:tileRect', attributes)
      end

      #
      # Write the <a:pattFill> element.
      #
      def write_a_patt_fill(pattern)
        attributes = [['prst', pattern[:pattern]]]

        @writer.tag_elements('a:pattFill', attributes) do
          write_a_fg_clr(pattern[:fg_color])
          write_a_bg_clr(pattern[:bg_color])
        end
      end

      def write_a_fg_clr(color)
        @writer.tag_elements('a:fgClr') { write_a_srgb_clr(color(color)) }
      end

      def write_a_bg_clr(color)
        @writer.tag_elements('a:bgClr') { write_a_srgb_clr(color(color)) }
      end

      #
      # Write the <c:numberFormat> element for data labels.
      #
      def write_data_label_number_format(format_code)
        source_linked = 0

        attributes = [
          ['formatCode',   format_code],
          ['sourceLinked', source_linked]
        ]

        @writer.empty_tag('c:numFmt', attributes)
      end

      #
      # Write the <c:dLbls> element.
      #
      def write_d_lbls(labels) # :nodoc:
        return unless labels

        @writer.tag_elements('c:dLbls') do
          # Write the custom c:dLbl elements.
          write_custom_labels(labels, labels[:custom]) if labels[:custom]
          # Write the c:numFmt element.
          write_data_label_number_format(labels[:num_format]) if labels[:num_format]
          # Write the c:spPr element.
          write_sp_pr(labels)
          # Write the data label font elements.
          write_axis_font(labels[:font]) if labels[:font]
          # Write the c:dLblPos element.
          write_d_lbl_pos(labels[:position]) if ptrue?(labels[:position])
          # Write the c:showLegendKey element.
          write_show_legend_key if labels[:legend_key]
          # Write the c:showVal element.
          write_show_val if labels[:value]
          # Write the c:showCatName element.
          write_show_cat_name if labels[:category]
          # Write the c:showSerName element.
          write_show_ser_name if labels[:series_name]
          # Write the c:showPercent element.
          write_show_percent if labels[:percentage]
          # Write the c:separator element.
          write_separator(labels[:separator]) if labels[:separator]
          # Write the c:showLeaderLines element.
          write_show_leader_lines if labels[:leader_lines]
        end
      end

      #
      # Write the <c:dLbl> element.
      #
      def write_custom_labels(parent, labels)
        index  = 0

        labels.each do |label|
          index += 1
          next unless ptrue?(label)

          @writer.tag_elements('c:dLbl') do
            # Write the c:idx element.
            write_idx(index - 1)

            if label[:delete]
              write_delete(1)
            elsif label[:formula]
              write_custom_label_formula(label)

              write_d_lbl_pos(parent[:position]) if parent[:position]
              write_show_val      if parent[:value]
              write_show_cat_name if parent[:category]
              write_show_ser_name if parent[:series_name]
            elsif label[:value]
              write_custom_label_str(label)

              write_d_lbl_pos(parent[:position]) if parent[:position]
              write_show_val      if parent[:value]
              write_show_cat_name if parent[:category]
              write_show_ser_name if parent[:series_name]
            else
              write_custom_label_format_only(label)
            end
          end
        end
      end

      #
      # Write parts of the <c:dLbl> element for strings.
      #
      def write_custom_label_str(label)
        value          = label[:value]
        font           = label[:font]
        is_y_axis      = 0
        has_formatting = has_fill_formatting(label)

        # Write the c:layout element.
        write_layout

        @writer.tag_elements('c:tx') do
          # Write the c:rich element.
          write_rich(value, font, is_y_axis, !has_formatting)
        end

        # Write the c:cpPr element.
        write_sp_pr(label)
      end

      #
      # Write parts of the <c:dLbl> element for formulas.
      #
      def write_custom_label_formula(label)
        formula = label[:formula]
        data_id = label[:data_id]

        data = @formula_data[data_id] if data_id

        # Write the c:layout element.
        write_layout

        @writer.tag_elements('c:tx') do
          # Write the c:strRef element.
          write_str_ref(formula, data, 'str')
        end

        # Write the data label formatting, if any.
        write_custom_label_format_only(label)
      end

      #
      # Write parts of the <c:dLbl> element for labels where only the formatting has
      # changed.
      #
      def write_custom_label_format_only(label)
        font           = label[:font]
        has_formatting = has_fill_formatting(label)

        if has_formatting
          # Write the c:spPr element.
          write_sp_pr(label)
          write_tx_pr(font)
        elsif font
          @writer.empty_tag('c:spPr')
          write_tx_pr(font)
        end
      end

      #
      # Write the <c:showLegendKey> element.
      #
      def write_show_legend_key
        @writer.empty_tag('c:showLegendKey', [['val', 1]])
      end

      #
      # Write the <c:showVal> element.
      #
      def write_show_val # :nodoc:
        @writer.empty_tag('c:showVal', [['val', 1]])
      end

      #
      # Write the <c:showCatName> element.
      #
      def write_show_cat_name # :nodoc:
        @writer.empty_tag('c:showCatName', [['val', 1]])
      end

      #
      # Write the <c:showSerName> element.
      #
      def write_show_ser_name # :nodoc:
        @writer.empty_tag('c:showSerName', [['val', 1]])
      end

      #
      # Write the <c:showPercent> element.
      #
      def write_show_percent
        @writer.empty_tag('c:showPercent', [['val', 1]])
      end

      #
      # Write the <c:separator> element.
      #
      def write_separator(data)
        @writer.data_element('c:separator', data)
      end

      # Write the <c:showLeaderLines> element. This is different for Pie/Doughnut
      # charts. Other chart types only supported leader lines after Excel 2015 via
      # an extension element.
      def write_show_leader_lines
        uri        = '{CE6537A1-D6FC-4f65-9D91-7224C49458BB}'
        xmlns_c_15 = 'http://schemas.microsoft.com/office/drawing/2012/chart'

        attributes1 = [
          ['uri', uri],
          ['xmlns:c15', xmlns_c_15]
        ]

        attributes2 = [['val',  1]]

        @writer.tag_elements('c:extLst') do
          @writer.tag_elements('c:ext', attributes1) do
            @writer.empty_tag('c15:showLeaderLines', attributes2)
          end
        end
      end

      #
      # Write the <c:dLblPos> element.
      #
      def write_d_lbl_pos(val)
        @writer.empty_tag('c:dLblPos', [['val', val]])
      end

      def has_fill_formatting(element)
        line     = series_property(element, :line)
        fill     = series_property(element, :fill)
        pattern  = series_property(element, :pattern)
        gradient = series_property(element, :gradient)

        (line && ptrue?(line[:_defined])) ||
          (fill && ptrue?(fill[:_defined])) || pattern || gradient
      end

      #
      # Write the <a:latin> element.
      #
      def write_a_latin(args) # :nodoc:
        @writer.empty_tag('a:latin', args)
      end
    end
  end
end
