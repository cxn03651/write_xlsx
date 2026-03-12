# -*- coding: utf-8 -*-
# frozen_string_literal: true

###############################################################################
#
# Chart
#
# The Chart class acts as a facade coordinating chart configuration,
# data management and XML generation.
#
# Responsibilities are delegated to specialized modules:
#
#   Initialization     - chart setup and defaults
#   Settings           - public chart configuration API
#   SeriesData         - series and formula bookkeeping
#   XmlWriter          - core chart XML assembly
#   AxisWriter         - axis related XML
#   SeriesWriter       - series XML generation
#   FormattingWriter   - formatting and label XML
#
###############################################################################

require 'write_xlsx/package/xml_writer_simple'
require 'write_xlsx/gradient'
require 'write_xlsx/chart/legend'
require 'write_xlsx/utility'
require 'write_xlsx/chart/axis'
require 'write_xlsx/chart/caption'
require 'write_xlsx/chart/series'
require 'write_xlsx/chart/chart_area'
require 'write_xlsx/chart/table'
require 'write_xlsx/chart/initialization'
require 'write_xlsx/chart/settings'
require 'write_xlsx/chart/series_data'
require 'write_xlsx/chart/xml_writer'
require 'write_xlsx/chart/axis_writer'
require 'write_xlsx/chart/series_writer'
require 'write_xlsx/chart/formatting_writer'

module Writexlsx
  class Chart
    include Writexlsx::Utility::Common
    include Writexlsx::Utility::ChartFormatting
    include Writexlsx::Utility::RichText
    include Writexlsx::Gradient
    include Initialization
    include Settings
    include SeriesData
    include XmlWriter
    include AxisWriter
    include SeriesWriter
    include FormattingWriter

    attr_accessor :id, :name                                         # :nodoc:
    attr_writer :index, :palette, :protection                        # :nodoc:
    attr_reader :embedded, :formula_ids, :formula_data               # :nodoc:
    attr_reader :x_scale, :y_scale, :x_offset, :y_offset             # :nodoc:
    attr_reader :width, :height                                      # :nodoc:
    attr_reader :label_positions, :label_position_default, :combined # :nodoc:
    attr_writer :date_category, :already_inserted                    # :nodoc:
    attr_writer :series_index                                        # :nodoc:
    attr_writer :writer                                              # :nodoc:
    attr_reader :x2_axis, :y2_axis, :axis2_ids                       # :nodoc:

    ###############################################################################
    #
    # Factory and lifecycle
    #
    ###############################################################################

    #
    # Factory method for returning chart objects based on their class type.
    #
    def self.factory(current_subclass, subtype = nil) # :nodoc:
      case current_subclass.downcase.capitalize
      when 'Area'
        require 'write_xlsx/chart/area'
        Chart::Area.new(subtype)
      when 'Bar'
        require 'write_xlsx/chart/bar'
        Chart::Bar.new(subtype)
      when 'Column'
        require 'write_xlsx/chart/column'
        Chart::Column.new(subtype)
      when 'Doughnut'
        require 'write_xlsx/chart/doughnut'
        Chart::Doughnut.new(subtype)
      when 'Line'
        require 'write_xlsx/chart/line'
        Chart::Line.new(subtype)
      when 'Pie'
        require 'write_xlsx/chart/pie'
        Chart::Pie.new(subtype)
      when 'Radar'
        require 'write_xlsx/chart/radar'
        Chart::Radar.new(subtype)
      when 'Scatter'
        require 'write_xlsx/chart/scatter'
        Chart::Scatter.new(subtype)
      when 'Stock'
        require 'write_xlsx/chart/stock'
        Chart::Stock.new(subtype)
      end
    end

    def initialize(subtype) # :nodoc:
      @writer = Package::XMLWriterSimple.new

      @subtype           = subtype
      @sheet_type        = 0x0200
      @series            = []
      @embedded          = false
      @id                = -1
      @series_index      = 0
      @style_id          = 2
      @formula_ids       = {}
      @formula_data      = []
      @protection        = 0
      @chartarea         = ChartArea.new
      @plotarea          = ChartArea.new
      @title             = Caption.new(self)
      @name              = ''
      @table             = nil
      set_default_properties
      @combined          = nil
      @is_secondary      = false
    end

    def set_xml_writer(filename) # :nodoc:
      @writer.set_xml_writer(filename)
    end

    ###############################################################################
    #
    # Chart type writing entry points
    #
    ###############################################################################

    #
    # Write the <c:barChart> element.
    #
    def write_bar_chart(params) # :nodoc:
      series = if ptrue?(params[:primary_axes])
                 get_primary_axes_series
               else
                 get_secondary_axes_series
               end
      return if series.empty?

      subtype = @subtype
      subtype = 'percentStacked' if subtype == 'percent_stacked'

      # Set a default overlap for stacked charts.
      @series_overlap_1 = 100 if @subtype =~ /stacked/ && !@series_overlap_1

      @writer.tag_elements('c:barChart') do
        # Write the c:barDir element.
        write_bar_dir
        # Write the c:grouping element.
        write_grouping(subtype)
        # Write the c:ser elements.
        series.each { |s| write_ser(s) }

        # write the c:marker element.
        write_marker_value

        if ptrue?(params[:primary_axes])
          # Write the c:gapWidth element.
          write_gap_width(@series_gap_1)
          # Write the c:overlap element.
          write_overlap(@series_overlap_1)
        else
          # Write the c:gapWidth element.
          write_gap_width(@series_gap_2)
          # Write the c:overlap element.
          write_overlap(@series_overlap_2)
        end

        # write the c:overlap element.
        write_overlap(@series_overlap)

        # Write the c:axId elements
        write_axis_ids(params)
      end
    end

    ###############################################################################
    #
    # Public chart state
    #
    ###############################################################################

    def already_inserted?
      @already_inserted
    end

    def is_secondary?
      @is_secondary
    end

    ###############################################################################
    #
    # private helpers
    #
    ###############################################################################
    private

    ###############################################################################
    #
    # Core chart XML helpers
    #
    ###############################################################################

    #
    # Write the chart type element. This method should be overridden by the
    # subclasses.
    #
    def write_chart_type; end

    #
    # Write the <c:protection> element.
    #
    def write_protection # :nodoc:
      return if @protection == 0

      @writer.empty_tag('c:protection')
    end

    ###############################################################################
    #
    # Extension list helpers
    #
    ###############################################################################

    def write_ext_lst_inverted_fill(color)
      uri = '{6F2FDCE9-48DA-4B69-8628-5D25D57E5C99}'
      xmlns_c_14 =
        'http://schemas.microsoft.com/office/drawing/2007/8/2/chart'

      attributes_1 = [
        ['uri', uri],
        ['xmlns:c14', xmlns_c_14]
      ]

      attributes_2 = [
        ['xmlns:c14', xmlns_c_14]
      ]

      @writer.tag_elements('c:extLst') do
        @writer.tag_elements('c:ext', attributes_1) do
          @writer.tag_elements('c14:invertSolidFillFmt') do
            @writer.tag_elements('c14:spPr', attributes_2) do
              write_a_solid_fill(color: color)
            end
          end
        end
      end
    end

    #
    # Write the <c:extLst> element for the display N/A as empty cell option.
    #
    def write_ext_lst_display_na
      uri        = '{56B9EC1D-385E-4148-901F-78D8002777C0}'
      xmlns_c_16 = 'http://schemas.microsoft.com/office/drawing/2017/03/chart'

      attributes1 = [
        ['uri', uri],
        ['xmlns:c16r3', xmlns_c_16]
      ]

      attributes2 = [
        ['val', 1]
      ]

      @writer.tag_elements('c:extLst') do
        @writer.tag_elements('c:ext', attributes1) do
          @writer.tag_elements('c16r3:dataDisplayOptions16') do
            @writer.empty_tag('c16r3:dispNaAsBlank', attributes2)
          end
        end
      end
    end

    ###############################################################################
    #
    # Legend and title helpers
    #
    ###############################################################################

    #
    # Write the <c:legend> element.
    #
    def write_legend # :nodoc:
      position = @legend.position.sub(/^overlay_/, '')
      return if position == 'none' || !position_allowed.has_key?(position)

      @delete_series = @legend.delete_series if @legend.delete_series.is_a?(Array)
      @writer.tag_elements('c:legend') do
        # Write the c:legendPos element.
        write_legend_pos(position_allowed[position])
        # Remove series labels from the legend.
        # Write the c:legendEntry element.
        @delete_series.each { |i| write_legend_entry(i) } if @delete_series
        # Write the c:layout element.
        write_layout(@legend.layout, 'legend')
        # Write the c:overlay element.
        write_overlay if @legend.position =~ /^overlay_/
        # Write the c:spPr element.
        write_sp_pr(@legend)
        # Write the c:txPr element.
        write_tx_pr(@legend.font) if ptrue?(@legend.font)
      end
    end

    def position_allowed
      {
        'right'     => 'r',
        'left'      => 'l',
        'top'       => 't',
        'bottom'    => 'b',
        'top_right' => 'tr'
      }
    end

    #
    # Write the <c:legendPos> element.
    #
    def write_legend_pos(val) # :nodoc:
      @writer.empty_tag('c:legendPos', [['val', val]])
    end

    #
    # Write the <c:legendEntry> element.
    #
    def write_legend_entry(index) # :nodoc:
      @writer.tag_elements('c:legendEntry') do
        # Write the c:idx element.
        write_idx(index)
        # Write the c:delete element.
        write_delete(1)
      end
    end

    #
    # Write the <c:overlay> element.
    #
    def write_overlay # :nodoc:
      @writer.empty_tag('c:overlay', [['val', 1]])
    end

    #
    # Write the <c:plotVisOnly> element.
    #
    def write_plot_vis_only # :nodoc:
      val = 1

      # Ignore this element if we are plotting hidden data.
      return if @show_hidden_data

      @writer.empty_tag('c:plotVisOnly', [['val', val]])
    end

    #
    # Write the <c:title> element for a rich string.
    #
    def write_title_rich(title, is_y_axis, font, layout, overlay = nil) # :nodoc:
      @writer.tag_elements('c:title') do
        # Write the c:tx element.
        write_tx_rich(title, is_y_axis, font)
        # Write the c:layout element.
        write_layout(layout, 'text')
        # Write the c:overlay element.
        write_overlay if overlay
      end
    end

    ###############################################################################
    #
    # Formatting and series decoration helpers
    #
    ###############################################################################

    #
    # Write the <c:dPt> elements.
    #
    def write_d_pt(points = nil)
      return unless ptrue?(points)

      index = -1
      points.each do |point|
        index += 1
        next unless ptrue?(point)

        write_d_pt_point(index, point)
      end
    end

    #
    # Write an individual <c:dPt> element.
    #
    def write_d_pt_point(index, point)
      @writer.tag_elements('c:dPt') do
        # Write the c:idx element.
        write_idx(index)
        # Write the c:spPr element.
        write_sp_pr(point)
      end
    end

    #
    # Write the <c:delete> element.
    #
    def write_delete(val) # :nodoc:
      @writer.empty_tag('c:delete', [['val', val]])
    end

    #
    # Write the <c:invertIfNegative> element.
    #
    def write_c_invert_if_negative(invert = nil) # :nodoc:
      return unless ptrue?(invert)

      @writer.empty_tag('c:invertIfNegative', [['val', 1]])
    end

    #
    # Write the <c:dTable> element.
    #
    def write_d_table
      @table.write_d_table(@writer) if @table
    end

    ###############################################################################
    #
    # Trendline helpers
    #
    ###############################################################################

    #
    # Write the <c:trendline> element.
    #
    def write_trendline(trendline) # :nodoc:
      return unless trendline

      @writer.tag_elements('c:trendline') do
        # Write the c:name element.
        write_name(trendline.name)
        # Write the c:spPr element.
        write_sp_pr(trendline)
        # Write the c:trendlineType element.
        write_trendline_type(trendline.type)
        # Write the c:order element for polynomial trendlines.
        write_trendline_order(trendline.order) if trendline.type == 'poly'
        # Write the c:period element for moving average trendlines.
        write_period(trendline.period) if trendline.type == 'movingAvg'
        # Write the c:forward element.
        write_forward(trendline.forward)
        # Write the c:backward element.
        write_backward(trendline.backward)
        if trendline.intercept
          # Write the c:intercept element.
          write_intercept(trendline.intercept)
        end
        if trendline.display_r_squared
          # Write the c:dispRSqr element.
          write_disp_rsqr
        end
        if trendline.display_equation
          # Write the c:dispEq element.
          write_disp_eq
          # Write the c:trendlineLbl element.
          write_trendline_lbl(trendline)
        end
      end
    end

    #
    # Write the <c:trendlineType> element.
    #
    def write_trendline_type(val) # :nodoc:
      @writer.empty_tag('c:trendlineType', [['val', val]])
    end

    #
    # Write the <c:name> element.
    #
    def write_name(data) # :nodoc:
      return unless data

      @writer.data_element('c:name', data)
    end

    #
    # Write the <c:order> element.
    #
    def write_trendline_order(val = 2) # :nodoc:
      @writer.empty_tag('c:order', [['val', val]])
    end

    #
    # Write the <c:period> element.
    #
    def write_period(val = 2) # :nodoc:
      @writer.empty_tag('c:period', [['val', val]])
    end

    #
    # Write the <c:forward> element.
    #
    def write_forward(val) # :nodoc:
      return unless val

      @writer.empty_tag('c:forward', [['val', val]])
    end

    #
    # Write the <c:backward> element.
    #
    def write_backward(val) # :nodoc:
      return unless val

      @writer.empty_tag('c:backward', [['val', val]])
    end

    #
    # Write the <c:intercept> element.
    #
    def write_intercept(val)
      @writer.empty_tag('c:intercept', [['val', val]])
    end

    #
    # Write the <c:dispEq> element.
    #
    def write_disp_eq
      @writer.empty_tag('c:dispEq', [['val', 1]])
    end

    #
    # Write the <c:dispRSqr> element.
    #
    def write_disp_rsqr
      @writer.empty_tag('c:dispRSqr', [['val', 1]])
    end

    #
    # Write the <c:trendlineLbl> element.
    #
    def write_trendline_lbl(trendline)
      @writer.tag_elements('c:trendlineLbl') do
        # Write the c:layout element.
        write_layout
        # Write the c:numFmt element.
        write_trendline_num_fmt
        # Write the c:spPr element for the label formatting.
        write_sp_pr(trendline.label)
        # Write the data label font elements.
        if trendline.label && ptrue?(trendline.label[:font])
          write_axis_font(trendline.label[:font])
        end
      end
    end

    #
    # Write the <c:numFmt> element.
    #
    def write_trendline_num_fmt
      format_code   = 'General'
      source_linked = 0

      attributes = [
        ['formatCode',   format_code],
        ['sourceLinked', source_linked]
      ]

      @writer.empty_tag('c:numFmt', attributes)
    end

    ###############################################################################
    #
    # Line and bar helpers
    #
    ###############################################################################

    #
    # Write the <c:hiLowLines> element.
    #
    def write_hi_low_lines # :nodoc:
      write_lines_base(@hi_low_lines, 'c:hiLowLines')
    end

    #
    # Write the <c:dropLines> elent.
    #
    def write_drop_lines
      write_lines_base(@drop_lines, 'c:dropLines')
    end

    def write_lines_base(lines, tag)
      return unless lines

      if lines.line_defined?
        @writer.tag_elements(tag) { write_sp_pr(lines) }
      else
        @writer.empty_tag(tag)
      end
    end

    #
    # Write the <c:upDownBars> element.
    #
    def write_up_down_bars
      return unless ptrue?(@up_down_bars)

      @writer.tag_elements('c:upDownBars') do
        # Write the c:gapWidth element.
        write_gap_width(150)

        # Write the c:upBars element.
        write_up_bars(@up_down_bars[:_up])

        # Write the c:downBars element.
        write_down_bars(@up_down_bars[:_down])
      end
    end

    #
    # Write the <c:gapWidth> element.
    #
    def write_gap_width(val = nil)
      return unless val

      @writer.empty_tag('c:gapWidth', [['val', val]])
    end

    #
    # Write the <c:upBars> element.
    #
    def write_up_bars(format)
      write_bars_base('c:upBars', format)
    end

    #
    # Write the <c:upBars> element.
    #
    def write_down_bars(format)
      write_bars_base('c:downBars', format)
    end

    #
    # Write the <c:smooth> element.
    #
    def write_c_smooth(smooth)
      return unless ptrue?(smooth)

      attributes = [['val', 1]]

      @writer.empty_tag('c:smooth', attributes)
    end

    def write_bars_base(tag, format)
      if format.line_defined? || format.fill_defined?
        @writer.tag_elements(tag) { write_sp_pr(format) }
      else
        @writer.empty_tag(tag)
      end
    end

    ###############################################################################
    #
    # Error bar helpers
    #
    ###############################################################################

    #
    # Write the X and Y error bars.
    #
    def write_error_bars(error_bars)
      return unless ptrue?(error_bars)

      write_err_bars('x', error_bars[:_x_error_bars]) if error_bars[:_x_error_bars]
      write_err_bars('y', error_bars[:_y_error_bars]) if error_bars[:_y_error_bars]
    end

    #
    # Write the <c:errBars> element.
    #
    def write_err_bars(direction, error_bars)
      return unless ptrue?(error_bars)

      @writer.tag_elements('c:errBars') do
        # Write the c:errDir element.
        write_err_dir(direction)

        # Write the c:errBarType element.
        write_err_bar_type(error_bars.direction)

        # Write the c:errValType element.
        write_err_val_type(error_bars.type)

        unless ptrue?(error_bars.endcap)
          # Write the c:noEndCap element.
          write_no_end_cap
        end

        case error_bars.type
        when 'stdErr'
          # Don't need to write a c:errValType tag.
        when 'cust'
          # Write the custom error tags.
          write_custom_error(error_bars)
        else
          # Write the c:val element.
          write_error_val(error_bars.value)
        end

        # Write the c:spPr element.
        write_sp_pr(error_bars)
      end
    end

    #
    # Write the <c:errDir> element.
    #
    def write_err_dir(val)
      @writer.empty_tag('c:errDir', [['val', val]])
    end

    #
    # Write the <c:errBarType> element.
    #
    def write_err_bar_type(val)
      @writer.empty_tag('c:errBarType', [['val', val]])
    end

    #
    # Write the <c:errValType> element.
    #
    def write_err_val_type(val)
      @writer.empty_tag('c:errValType', [['val', val]])
    end

    #
    # Write the <c:noEndCap> element.
    #
    def write_no_end_cap
      @writer.empty_tag('c:noEndCap', [['val', 1]])
    end

    #
    # Write the <c:val> element.
    #
    def write_error_val(val)
      @writer.empty_tag('c:val', [['val', val]])
    end

    #
    # Write the custom error bars type.
    #
    def write_custom_error(error_bars)
      if ptrue?(error_bars.plus_values)
        write_custom_error_base('c:plus',  error_bars.plus_values,  error_bars.plus_data)
        write_custom_error_base('c:minus', error_bars.minus_values, error_bars.minus_data)
      end
    end

    def write_custom_error_base(tag, values, data)
      @writer.tag_elements(tag) do
        write_num_ref_or_lit(values, data)
      end
    end

    def write_num_ref_or_lit(values, data)
      if values.to_s =~ /^=/                # '=Sheet1!$A$1:$A$5'
        write_num_ref(values, data, 'num')
      else                                  # [1, 2, 3]
        write_num_lit(values)
      end
    end
  end
end
