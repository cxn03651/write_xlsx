# -*- coding: utf-8 -*-
# frozen_string_literal: true

module Writexlsx
  class Workbook
    #
    # Workbook initialization helpers extracted from Workbook to keep the main
    # class focused on public API and orchestration.
    #
    module Initialization
      private

      def setup_core_state(file, options, default_formats)
        @options = options.dup                 # for test
        @default_formats = default_formats.dup # for test
        @writer = Package::XMLWriterSimple.new

        @file = file
        @tempdir = options[:tempdir] ||
                   File.join(
                     Dir.tmpdir,
                     Digest::MD5.hexdigest("#{Time.now.to_f}-#{Process.pid}")
                   )

        @date_1904       = options[:date_1904] || false
        @activesheet     = 0
        @firstsheet      = 0
        @selected        = 0
        @fileclosed      = false
        @optimization    = options[:optimization] || 0
        @excel2003_style = options[:excel2003_style] || false
        @read_only       = 0

        @strings_to_urls =
          options[:strings_to_urls].nil? || options[:strings_to_urls] ? true : false

        @max_url_length = if options[:max_url_length]
                            [options[:max_url_length].to_i, 255].max
                          else
                            MAX_URL_LENGTH
                          end
      end

      def setup_workbook_state(_options)
        @worksheets        = Sheets.new
        @charts            = []
        @drawings          = []
        @defined_names     = []
        @named_ranges      = []
        @custom_colors     = []
        @doc_properties    = {}
        @custom_properties = []
        @image_types       = {}
        @images            = []
        @has_comments      = false
        @has_metadata      = false
        @has_embedded_images = false
        @has_embedded_descriptions = false

        @x_window      = 240
        @y_window      = 15
        @window_width  = 16_095
        @window_height = 9_660
        @tab_ratio     = 600
      end

      def setup_format_state(_default_formats)
        @formats     = Formats.new
        @xf_formats  = []
        @dxf_formats = []
        @num_formats = []
      end

      def setup_shared_strings
        @shared_strings = Package::SharedStrings.new
      end

      def setup_embedded_assets
        @embedded_image_indexes = {}
        @embedded_images        = []
      end

      def setup_calculation_state
        @calc_id      = 124519
        @calc_mode    = 'auto'
        @calc_on_load = true
      end

      def setup_default_formats
        if @excel2003_style
          add_format(
            @default_formats.merge(
              xf_index:    0,
              font_family: 0,
              font:        'Arial',
              size:        10,
              theme:       -1
            )
          )
        else
          add_format(@default_formats.merge(xf_index: 0))
        end

        # Add a default URL format.
        @default_url_format = add_format(hyperlink: 1)
      end

      #
      # Workbook の生成時のオプションハッシュを解析する
      #
      def process_workbook_options(*params)
        case params.size
        when 0
          [{}, {}]
        when 1 # one hash
          options_keys = %i[tempdir date_1904 optimization excel2003_style strings_to_urls max_url_length]

          hash = params.first
          options = hash.select { |k, _v| options_keys.include?(k) }

          default_format_properties =
            hash[:default_format_properties] ||
            hash.reject { |k, _v| options_keys.include?(k) }

          [options, default_format_properties.dup]
        when 2 # array which includes options and default_format_properties
          options, default_format_properties = params
          default_format_properties ||= {}

          [options.dup, default_format_properties.dup]
        end
      end

      def filename
        setup_filename unless @filename
        @filename
      end

      def fileobj
        setup_filename unless @fileobj
        @fileobj
      end

      def setup_filename # :nodoc:
        if @file.respond_to?(:to_str) && @file != ''
          @filename = @file
          @fileobj  = nil
        elsif @file.respond_to?(:write)
          @filename = File.join(tempdir, Digest::MD5.hexdigest(Time.now.to_s) + '.xlsx.tmp')
          @fileobj  = @file
        else
          raise "'#{@file}' must be valid filename String of IO object."
        end
      end

      attr_reader :tempdir

      #
      # Sets the colour palette to the Excel defaults.
      #
      def set_color_palette # :nodoc:
        @palette = [
          [0x00, 0x00, 0x00, 0x00],    # 8
          [0xff, 0xff, 0xff, 0x00],    # 9
          [0xff, 0x00, 0x00, 0x00],    # 10
          [0x00, 0xff, 0x00, 0x00],    # 11
          [0x00, 0x00, 0xff, 0x00],    # 12
          [0xff, 0xff, 0x00, 0x00],    # 13
          [0xff, 0x00, 0xff, 0x00],    # 14
          [0x00, 0xff, 0xff, 0x00],    # 15
          [0x80, 0x00, 0x00, 0x00],    # 16
          [0x00, 0x80, 0x00, 0x00],    # 17
          [0x00, 0x00, 0x80, 0x00],    # 18
          [0x80, 0x80, 0x00, 0x00],    # 19
          [0x80, 0x00, 0x80, 0x00],    # 20
          [0x00, 0x80, 0x80, 0x00],    # 21
          [0xc0, 0xc0, 0xc0, 0x00],    # 22
          [0x80, 0x80, 0x80, 0x00],    # 23
          [0x99, 0x99, 0xff, 0x00],    # 24
          [0x99, 0x33, 0x66, 0x00],    # 25
          [0xff, 0xff, 0xcc, 0x00],    # 26
          [0xcc, 0xff, 0xff, 0x00],    # 27
          [0x66, 0x00, 0x66, 0x00],    # 28
          [0xff, 0x80, 0x80, 0x00],    # 29
          [0x00, 0x66, 0xcc, 0x00],    # 30
          [0xcc, 0xcc, 0xff, 0x00],    # 31
          [0x00, 0x00, 0x80, 0x00],    # 32
          [0xff, 0x00, 0xff, 0x00],    # 33
          [0xff, 0xff, 0x00, 0x00],    # 34
          [0x00, 0xff, 0xff, 0x00],    # 35
          [0x80, 0x00, 0x80, 0x00],    # 36
          [0x80, 0x00, 0x00, 0x00],    # 37
          [0x00, 0x80, 0x80, 0x00],    # 38
          [0x00, 0x00, 0xff, 0x00],    # 39
          [0x00, 0xcc, 0xff, 0x00],    # 40
          [0xcc, 0xff, 0xff, 0x00],    # 41
          [0xcc, 0xff, 0xcc, 0x00],    # 42
          [0xff, 0xff, 0x99, 0x00],    # 43
          [0x99, 0xcc, 0xff, 0x00],    # 44
          [0xff, 0x99, 0xcc, 0x00],    # 45
          [0xcc, 0x99, 0xff, 0x00],    # 46
          [0xff, 0xcc, 0x99, 0x00],    # 47
          [0x33, 0x66, 0xff, 0x00],    # 48
          [0x33, 0xcc, 0xcc, 0x00],    # 49
          [0x99, 0xcc, 0x00, 0x00],    # 50
          [0xff, 0xcc, 0x00, 0x00],    # 51
          [0xff, 0x99, 0x00, 0x00],    # 52
          [0xff, 0x66, 0x00, 0x00],    # 53
          [0x66, 0x66, 0x99, 0x00],    # 54
          [0x96, 0x96, 0x96, 0x00],    # 55
          [0x00, 0x33, 0x66, 0x00],    # 56
          [0x33, 0x99, 0x66, 0x00],    # 57
          [0x00, 0x33, 0x00, 0x00],    # 58
          [0x33, 0x33, 0x00, 0x00],    # 59
          [0x99, 0x33, 0x00, 0x00],    # 60
          [0x99, 0x33, 0x66, 0x00],    # 61
          [0x33, 0x33, 0x99, 0x00],    # 62
          [0x33, 0x33, 0x33, 0x00]    # 63
        ]
      end
    end
  end
end
