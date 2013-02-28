# -*- coding: utf-8 -*-

module Writexlsx
  ###############################################################################
  #
  # Shape - A class for writing Excel shapes.
  #
  # Used in conjunction with Excel::Writer::XLSX.
  #
  # Copyright 2000-2012, John McNamara, jmcnamara@cpan.org
  # Converted to ruby by Hideo NAKAMURA, cxn03651@msj.biglobe.ne.jp
  #
  class Shape

    attr_reader :edit_as, :type, :drawing
    attr_reader :tx_box, :fill, :line, :format
    attr_reader :align, :valign
    attr_accessor :name, :connect, :type, :id, :start, :end, :rotation
    attr_accessor :flip_h, :flip_v, :adjustments, :palette, :text, :stencil
    attr_accessor :row_start, :row_end, :column_start, :column_end
    attr_accessor :x1, :x2, :y1, :y2, :x_abs, :y_abs, :start_index, :end_index
    attr_accessor :x_offset, :y_offset, :width, :height, :scale_x, :scale_y
    attr_accessor :width_emu, :height_emu, :element, :line_weight, :line_type
    attr_accessor :start_side, :end_side

    def initialize(properties = {})
      @writer = Package::XMLWriterSimple.new
      @name   = nil
      @type   = 'rect'

      # Is a Connector shape. 1/0 Value is a hash lookup from type.
      @connect = 0

      # Is a Drawing. Always 0, since a single shape never fills an entire sheet.
      @drawing = 0

      # OneCell or Absolute: options to move and/or size with cells.
      @edit_as = nil

      # Auto-incremented, unless supplied by user.
      @id = 0

      # Shape text (usually centered on shape geometry).
      @text = 0

      # Shape stencil mode.  A copy (child) is created when inserted.
      # The link to parent is broken.
      @stencil = 1

      # Index to _shapes array when inserted.
      @element = -1

      # Shape ID of starting connection, if any.
      @start = nil

      # Shape vertex, starts at 0, numbered clockwise from 12 o'clock.
      @start_index = nil

      @end       = nil
      @end_index = nil

      # Number and size of adjustments for shapes (usually connectors).
      @adjustments = []

      # Start and end sides. t)op, b)ottom, l)eft, or r)ight.
      @start_side = ''
      @end_side   = ''

      # Flip shape Horizontally. eg. arrow left to arrow right.
      @flip_h = 0

      # Flip shape Vertically. eg. up arrow to down arrow.
      @flip_v = 0

      # shape rotation (in degrees 0-360).
      @rotation = 0

      # An alternate way to create a text box, because Excel allows it.
      # It is just a rectangle with text.
      @tx_box = 0

      # Shape outline colour, or 0 for noFill (default black).
      @line = '000000'

      # Line type: dash, sysDot, dashDot, lgDash, lgDashDot, lgDashDotDot.
      @line_type = ''

      # Line weight (integer).
      @line_weight = 1

      # Shape fill colour, or 0 for noFill (default noFill).
      @fill = 0

      # Formatting for shape text, if any.
      @format = {}

      # copy of colour palette table from Workbook.pm.
      @palette = []

      # Vertical alignment: t, ctr, b.
      @valign = 'ctr'

      # Alignment: l, ctr, r, just
      @align = 'ctr'

      @x_offset = 0
      @y_offset = 0

      # Scale factors, which also may be set when the shape is inserted.
      @scale_x = 1
      @scale_y = 1

      # Default size, which can be modified and/or scaled.
      @width  = 50
      @height = 50

      # Initial assignment. May be modified when prepared.
      @column_start = 0
      @row_start    = 0
      @x1           = 0
      @y1           = 0
      @column_end   = 0
      @row_end      = 0
      @x2           = 0
      @y2           = 0
      @x_abs        = 0
      @y_abs        = 0

      set_properties(properties)
    end

    def set_properties(properties)
      # Override default properties with passed arguments
      properties.each do |key, value|
        # Strip leading "-" from Tk style properties e.g. -color => 'red'.
        k = key.to_s.sub(/^-/, '')
        self.instance_variable_set("@#{key}", value)
=begin
           if key.to_s == 'format'
          @format = value
        elsif value.respond_to?(:coerce)
          eval "@#{k} = #{value}"
        else
          eval "@#{k} = %!#{value}!"
        end
=end
      end
    end

    #
    # Set the shape adjustments array (as a reference).
    #
    def adjustments=(args)
      @adjustments = *args
    end
=begin
    def [](attr)
      self.instance_variable_get("@#{attr}")
    end

    def []=(attr, value)
      self.instance_variable_set("@#{attr}", value)
    end
=end
    #
    # Convert from an Excel internal colour index to a XML style #RRGGBB index
    # based on the default or user defined values in the Workbook palette.
    # Note: This version doesn't add an alpha channel.
    #
    def get_palette_color(index)
      # Adjust the colour index.
      idx = index - 8

      # Palette is passed in from the Workbook class.
      rgb = @palette[idx]

      sprintf("%02X%02X%02X", *rgb)
    end
  end
end
