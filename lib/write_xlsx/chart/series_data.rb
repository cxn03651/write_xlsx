# -*- coding: utf-8 -*-
# frozen_string_literal: true

###############################################################################
#
# series_data.rb - series/formula bookkeeping and axis id management
#
###############################################################################

module Writexlsx
  class Chart
    module SeriesData
      #
      # Switch name and name_formula parameters if required.
      #
      def process_names(name = nil, name_formula = nil) # :nodoc:
        # Name looks like a formula, use it to set name_formula.
        if name.respond_to?(:to_ary)
          cell = xl_rowcol_to_cell(name[1], name[2], 1, 1)
          name_formula = "#{quote_sheetname(name[0])}!#{cell}"
          name = ''
        elsif name && name =~ /^=[^!]+!\$/
          name_formula = name
          name         = ''
        end

        [name, name_formula]
      end

      #
      # Assign an id to a each unique series formula or title/axis formula. Repeated
      # formulas such as for categories get the same id. If the series or title
      # has user specified data associated with it then that is also stored. This
      # data is used to populate cached Excel data when creating a chart.
      # If there is no user defined data then it will be populated by the parent
      # workbook in Workbook::_add_chart_data
      #
      def data_id(full_formula, data) # :nodoc:
        return unless full_formula

        # Strip the leading '=' from the formula.
        formula = full_formula.sub(/^=/, '')

        # Store the data id in a hash keyed by the formula and store the data
        # in a separate array with the same id.
        if @formula_ids.has_key?(formula)
          # Formula already seen. Return existing id.
          id = @formula_ids[formula]
          # Store user defined data if it isn't already there.
          @formula_data[id] ||= data
        else
          # Haven't seen this formula before.
          id = @formula_ids[formula] = @formula_data.size
          @formula_data << data
        end

        id
      end

      private

      #
      # retun primary/secondary series by :primary_axes flag
      #
      def axes_series(params)
        if params[:primary_axes] == 0
          secondary_axes_series
        else
          primary_axes_series
        end
      end

      #
      # Find the overall type of the data associated with a series.
      #
      # TODO. Need to handle date type.
      #
      def get_data_type(data) # :nodoc:
        # Check for no data in the series.
        return 'none' unless data
        return 'none' if data.empty?
        return 'multi_str' if data.first.is_a?(Array)

        # If the token isn't a number assume it is a string.
        data.each do |token|
          next unless token
          return 'str' unless token.is_a?(Numeric)
        end

        # The series data was all numeric.
        'num'
      end

      #
      # Returns series which use the primary axes.
      #
      def get_primary_axes_series
        @series.reject(&:y2_axis)
      end
      alias primary_axes_series get_primary_axes_series

      #
      # Returns series which use the secondary axes.
      #
      def get_secondary_axes_series
        @series.select(&:y2_axis)
      end
      alias secondary_axes_series get_secondary_axes_series

      #
      # Add a unique ids for primary or secondary axis.
      #
      def add_axis_ids(params) # :nodoc:
        if ptrue?(params[:primary_axes])
          @axis_ids  += ids
        else
          @axis2_ids += ids
        end
      end

      def ids
        chart_id   = 5001 + @id
        axis_count = 1 + @axis2_ids.size + @axis_ids.size

        id1 = sprintf('%04d%04d', chart_id, axis_count)
        id2 = sprintf('%04d%04d', chart_id, axis_count + 1)

        [id1, id2]
      end
    end
  end
end
