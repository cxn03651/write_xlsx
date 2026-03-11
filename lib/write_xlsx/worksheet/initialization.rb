# -*- coding: utf-8 -*-
# frozen_string_literal: true

module Writexlsx
  class Worksheet
    module Initialization
      def setup_identity(workbook, index, name)
        @workbook = workbook
        @index = index
        @name = name

        @excel_version = 2007
        @palette = workbook.palette
        @default_url_format = workbook.default_url_format
        @max_url_length = workbook.max_url_length
      end

      def setup_limits
        @xls_rowmax = 1_048_576
        @xls_colmax = 16_384
        @xls_strmax = 32_767

        @dim_rowmin = nil
        @dim_rowmax = nil
        @dim_colmin = nil
        @dim_colmax = nil
      end

      def setup_dependencies
        @page_setup = Writexlsx::PageSetup.new
        @comments   = Package::Comments.new(self)
        @assets     = AssetManager.new
      end

      def setup_view_options
        @screen_gridlines = true
        @show_zeros = true
        @hide_row_col_headers = 0
        @top_left_cell = ''

        @tab_color = 0

        @zoom = 100
        @zoom_scale_normal = true
        @right_to_left = false
        @leading_zeros = false
      end

      def setup_sheet_geometry
        @outline_row_level = 0
        @outline_col_level = 0

        @original_row_height = 15
        @default_row_height = 15
        @default_row_pixels = 20
        @default_col_width = 8.43
        @default_date_pixels = 68
      end

      def setup_row_and_column_state
        @col_info = {}
        @cell_data_store = CellDataStore.new

        @set_cols = {}
        @set_rows = {}
        @row_sizes = {}

        @col_size_changed = false
      end

      def setup_filter_and_selection_state
        @selections = []
        @panes = []

        @autofilter_area = nil
        @filter_on = false
        @filter_range = []
        @filter_cols = {}
        @filter_cells = {}
        @filter_type = {}
      end

      def setup_drawing_and_media
        @last_shape_id = 1
        @rel_count = 0

        @external_hyper_links = []
        @external_drawing_links = []
        @external_comment_links = []
        @external_vml_links = []
        @external_background_links = []
        @external_table_links = []

        @drawing_links = []
        @vml_drawing_links = []

        @shape_hash = {}
        @drawing_rels = {}
        @drawing_rels_id = 0
        @vml_drawing_rels = {}
        @vml_drawing_rels_id = 0

        @has_dynamic_functions = false
        @has_embedded_images = false
        @use_future_functions = false
        @has_vml = false

        @buttons_array = []
        @header_images_array = []
      end

      def setup_cell_features
        @merge = []

        @validations = []
        @cond_formats = {}
        @data_bars_2010 = []
        @dxf_priority = 1

        @ignore_errors = nil
      end

      def setup_protection
        @protected_ranges = []
        @num_protected_ranges = 0
      end

      def apply_excel2003_compatibility
        @original_row_height = 12.75
        @default_row_height = 12.75
        @default_row_pixels = 17

        self.margins_left_right = 0.75
        self.margins_top_bottom = 1

        @page_setup.margin_header = 0.5
        @page_setup.margin_footer = 0.5
        @page_setup.header_footer_aligns = false
      end

      def setup_workbook_dependent_state
        @embedded_image_indexes = @workbook.embedded_image_indexes
      end
    end
  end
end
