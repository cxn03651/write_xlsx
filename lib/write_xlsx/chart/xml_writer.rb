# -*- coding: utf-8 -*-
# frozen_string_literal: true

###############################################################################
#
# xml_writer.rb - core chart XML assembly flow
#
###############################################################################

module Writexlsx
  class Chart
    module XmlWriter
      #
      # Assemble and write the XML file.
      #
      def assemble_xml_file   # :nodoc:
        write_xml_declaration do
          # Write the c:chartSpace element.
          write_chart_space do
            # Write the c:lang element.
            write_lang
            # Write the c:style element.
            write_style
            # Write the c:protection element.
            write_protection
            # Write the c:chart element.
            write_chart
            # Write the c:spPr element for the chartarea formatting.
            write_sp_pr(@chartarea)
            # Write the c:printSettings element.
            write_print_settings if @embedded
          end
        end
      end

      private

      #
      # Write the <c:chartSpace> element.
      #
      def write_chart_space(&block) # :nodoc:
        @writer.tag_elements('c:chartSpace', chart_space_attributes, &block)
      end

      # for <c:chartSpace> element.
      def chart_space_attributes # :nodoc:
        schema  = 'http://schemas.openxmlformats.org/'
        [
          ['xmlns:c', "#{schema}drawingml/2006/chart"],
          ['xmlns:a', "#{schema}drawingml/2006/main"],
          ['xmlns:r', "#{schema}officeDocument/2006/relationships"]
        ]
      end

      #
      # Write the <c:lang> element.
      #
      def write_lang # :nodoc:
        @writer.empty_tag('c:lang', [%w[val en-US]])
      end

      #
      # Write the <c:style> element.
      #
      def write_style # :nodoc:
        return if @style_id == 2

        @writer.empty_tag('c:style', [['val', @style_id]])
      end

      #
      # Write the <c:chart> element.
      #
      def write_chart # :nodoc:
        @writer.tag_elements('c:chart') do
          # Write the chart title elements.
          if @title.none
            # Turn off the title.
            write_auto_title_deleted
          elsif @title.formula
            write_title_formula(@title, nil, nil, @title.layout, @title.overlay)
          elsif @title.name
            write_title_rich(@title, nil, @title.name_font, @title.layout, @title.overlay)
          end

          # Write the c:plotArea element.
          write_plot_area
          # Write the c:legend element.
          write_legend
          # Write the c:plotVisOnly element.
          write_plot_vis_only

          # Write the c:dispBlanksAs element.
          write_disp_blanks_as

          # Write the c:extLst element.
          write_ext_lst_display_na if @show_na_as_empty
        end
      end

      #
      # Write the <c:plotArea> element.
      #
      # This method orchestrates the plot area generation including:
      #   - chart type elements
      #   - primary and secondary axes
      #   - combined chart handling
      #   - plot area formatting
      #
      # The logic is intentionally split into small helpers to preserve the
      # complex ordering requirements of Excel chart XML.
      #
      def write_plot_area # :nodoc:
        second_chart = @combined

        @writer.tag_elements('c:plotArea') do
          write_plot_area_layout
          write_plot_area_chart_types(second_chart)
          write_plot_area_primary_axes
          write_plot_area_secondary_axes(second_chart)
          write_plot_area_formatting
        end
      end

      def write_plot_area_layout
        write_layout(@plotarea.layout, 'plot')
      end

      def write_plot_area_chart_types(second_chart)
        write_primary_and_secondary_chart_types
        write_combined_chart_types(second_chart) if second_chart
      end

      def write_primary_and_secondary_chart_types
        write_chart_type(primary_axes: 1)
        write_chart_type(primary_axes: 0)
      end

      def write_combined_chart_types(second_chart)
        prepare_combined_chart(second_chart)
        second_chart.write_chart_type(primary_axes: 1)
        second_chart.write_chart_type(primary_axes: 0)
      end

      def prepare_combined_chart(second_chart)
        second_chart.id = second_chart.is_secondary? ? 1000 + @id : @id
        second_chart.writer = @writer
        second_chart.series_index = @series_index
      end

      def write_plot_area_primary_axes
        params = {
          x_axis:   @x_axis,
          y_axis:   @y_axis,
          axis_ids: @axis_ids
        }

        write_category_axis_for(params)
        write_val_axis(@x_axis, @y_axis, @axis_ids)
      end

      def write_plot_area_secondary_axes(second_chart)
        params = {
          x_axis:   @x2_axis,
          y_axis:   @y2_axis,
          axis_ids: @axis2_ids
        }

        write_val_axis(@x2_axis, @y2_axis, @axis2_ids)

        if second_chart && second_chart.is_secondary?
          params = {
            x_axis:   second_chart.x2_axis,
            y_axis:   second_chart.y2_axis,
            axis_ids: second_chart.axis2_ids
          }

          second_chart.write_val_axis(
            second_chart.x2_axis,
            second_chart.y2_axis,
            second_chart.axis2_ids
          )
        end

        write_category_axis_for(params)
      end

      def write_category_axis_for(params)
        if @date_category
          write_date_axis(params)
        else
          write_cat_axis(params)
        end
      end

      def write_plot_area_formatting
        write_d_table
        write_sp_pr(@plotarea)
      end

      #
      # Write the <c:dispBlanksAs> element.
      #
      def write_disp_blanks_as
        return if @show_blanks == 'gap'

        @writer.empty_tag('c:dispBlanksAs', [['val', @show_blanks]])
      end

      #
      # Write the <c:layout> element.
      #
      def write_layout(layout = nil, type = nil) # :nodoc:
        tag = 'c:layout'

        if layout
          @writer.tag_elements(tag)  { write_manual_layout(layout, type) }
        else
          @writer.empty_tag(tag)
        end
      end

      #
      # Write the <c:manualLayout> element.
      #
      def write_manual_layout(layout, type)
        @writer.tag_elements('c:manualLayout') do
          # Plotarea has a layoutTarget element.
          @writer.empty_tag('c:layoutTarget', [%w[val inner]]) if type == 'plot'

          # Set the x, y positions.
          @writer.empty_tag('c:xMode', [%w[val edge]])
          @writer.empty_tag('c:yMode', [%w[val edge]])
          @writer.empty_tag('c:x',     [['val', layout[:x]]])
          @writer.empty_tag('c:y',     [['val', layout[:y]]])

          # For plotarea and legend set the width and height.
          if type != 'text'
            @writer.empty_tag('c:w', [['val', layout[:width]]])
            @writer.empty_tag('c:h', [['val', layout[:height]]])
          end
        end
      end

      #
      # Write the <c:printSettings> element.
      #
      def write_print_settings # :nodoc:
        @writer.tag_elements('c:printSettings') do
          # Write the c:headerFooter element.
          write_header_footer
          # Write the c:pageMargins element.
          write_page_margins
          # Write the c:pageSetup element.
          write_page_setup
        end
      end

      #
      # Write the <c:headerFooter> element.
      #
      def write_header_footer # :nodoc:
        @writer.empty_tag('c:headerFooter')
      end

      #
      # Write the <c:pageMargins> element.
      #
      def write_page_margins # :nodoc:
        attributes = [
          ['b',      0.75],
          ['l',      0.7],
          ['r',      0.7],
          ['t',      0.75],
          ['header', 0.3],
          ['footer', 0.3]
        ]

        @writer.empty_tag('c:pageMargins', attributes)
      end

      #
      # Write the <c:pageSetup> element.
      #
      def write_page_setup # :nodoc:
        @writer.empty_tag('c:pageSetup')
      end

      #
      # Write the <c:autoTitleDeleted> element.
      #
      def write_auto_title_deleted
        attributes = [['val', 1]]

        @writer.empty_tag('c:autoTitleDeleted', attributes)
      end

      #
      # Write the <c:grouping> element.
      #
      def write_grouping(val) # :nodoc:
        @writer.empty_tag('c:grouping', [['val', val]])
      end

      #
      # Write the <c:overlap> element.
      #
      def write_overlap(val = nil) # :nodoc:
        return unless val

        @writer.empty_tag('c:overlap', [['val', val]])
      end
    end
  end
end
