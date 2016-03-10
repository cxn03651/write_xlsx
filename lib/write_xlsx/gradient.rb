module Writexlsx
  module Gradient
    def gradient_properties(args)
      return unless ptrue?(args)

      gradient = {}

      types    = {
        'linear'      => 'linear',
        'radial'      => 'circle',
        'rectangular' => 'rect',
        'path'        => 'shape'
      }

      # Check the colors array exists and is valid.
      raise "Gradient must include colors array" unless ptrue?(args[:colors])
      # Check the colors array has the right number of entries.
      raise "Gradient colors array must include at least 2 values" if args[:colors].size < 2

      gradient[:colors] = args[:colors]

      if ptrue?(args[:positions])
        # Check the positions array has the right number of entries.
        raise "Gradient positions not equal to numbers of colors" unless args[:positions].size == args[:colors].size

        # Check the positions are in the correct range.
        args[:positions].each do |pos|
          if pos < 0 || pos > 100
            raise "Gradient position '#{pos} must be in range 0 <= pos <= 100"
          end
        end
        gradient[:positions] = args[:positions]
      else
        # Use the default gradient positions.
        case args[:colors].size
        when 2
          gradient[:positions] = [0, 100]
        when 3
          gradient[:positions] = [0, 50, 100]
        when 4
          gradient[:positions] = [0, 33, 66, 100]
        else
          raise "Must specify gradient positions"
        end
      end

      # Set the gradient angle.
      if args[:angle]
        angle = args[:angle]

        if angle < 0 || angle > 359.9
          raise "Gradient angle '#{angle} must be in range 0 <= pos < 360"
        end
        gradient[:angle] = angle
      else
        gradient[:angle] = 90
      end

      # Set the gradient type.
      if args[:type]
        type = args[:type]

        unless types[type]
          raise "Unknow gradient type '#{type}'"
        end
        gradient[:type] = types[type]
      else
        gradient[:type] = 'linear'
      end

      gradient
    end
  end
end
