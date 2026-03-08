# frozen_string_literal: true

module Writexlsx
  class Worksheet
    # Protection-related operations extracted from Worksheet to slim the main class.
    module Protection
      #
      # Set the worksheet protection flags to prevent modification of worksheet
      # objects.
      #
      def protect(password = nil, options = {})
        check_parameter(options, protect_default_settings.keys, 'protect')
        @protect = protect_default_settings.merge(options)

        # Set the password after the user defined values.
        if password && password != ''
          @protect[:password] =
            encode_password(password)
        end
      end

      #
      # Unprotect ranges within a protected worksheet.
      #
      def unprotect_range(range, range_name = nil, password = nil)
        if range.nil?
          raise "The range must be defined in unprotect_range())\n"
        else
          range = range.gsub("$", "")
          range = range.sub(/^=/, "")
          @num_protected_ranges += 1
        end

        range_name ||= "Range#{@num_protected_ranges}"
        password   &&= encode_password(password)

        @protected_ranges << [range, range_name, password]
      end

      protected

      def protect_default_settings  # :nodoc:
        {
          sheet:                 true,
          content:               false,
          objects:               false,
          scenarios:             false,
          format_cells:          false,
          format_columns:        false,
          format_rows:           false,
          insert_columns:        false,
          insert_rows:           false,
          insert_hyperlinks:     false,
          delete_columns:        false,
          delete_rows:           false,
          select_locked_cells:   true,
          sort:                  false,
          autofilter:            false,
          pivot_tables:          false,
          select_unlocked_cells: true
        }
      end
    end
  end
end
