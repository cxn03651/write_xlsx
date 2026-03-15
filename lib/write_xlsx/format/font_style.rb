# -*- encoding: utf-8 -*-
# frozen_string_literal: true

module Writexlsx
  class Format
    class FontStyle
      def initialize(format)
        @format = format
      end

      def index
        @format.instance_variable_get(:@font_state).index
      end

      def index=(value)
        @format.instance_variable_get(:@font_state).index = value
        @format.send(:sync_font_ivars_from_state)
      end

      def name
        @format.instance_variable_get(:@font_state).name
      end

      def name=(value)
        @format.instance_variable_get(:@font_state).name = value
        @format.send(:sync_font_ivars_from_state)
      end

      def size
        @format.instance_variable_get(:@font_state).size
      end

      def size=(value)
        @format.instance_variable_get(:@font_state).size = value
        @format.send(:sync_font_ivars_from_state)
      end

      def bold
        @format.instance_variable_get(:@font_state).bold
      end

      def bold=(value)
        @format.instance_variable_get(:@font_state).bold = value
        @format.send(:sync_font_ivars_from_state)
      end

      def italic
        @format.instance_variable_get(:@font_state).italic
      end

      def italic=(value)
        @format.instance_variable_get(:@font_state).italic = value
        @format.send(:sync_font_ivars_from_state)
      end

      def color
        @format.instance_variable_get(:@font_state).color
      end

      def color=(value)
        @format.instance_variable_get(:@font_state).color = value
        @format.send(:sync_font_ivars_from_state)
      end

      def color_indexed
        @format.instance_variable_get(:@font_state).color_indexed
      end

      def color_indexed=(value)
        @format.instance_variable_get(:@font_state).color_indexed = value
        @format.send(:sync_font_ivars_from_state)
      end

      def underline
        @format.instance_variable_get(:@font_state).underline
      end

      def underline=(value)
        @format.instance_variable_get(:@font_state).underline = value
        @format.send(:sync_font_ivars_from_state)
      end

      def strikeout
        @format.instance_variable_get(:@font_state).strikeout
      end

      def strikeout=(value)
        @format.instance_variable_get(:@font_state).strikeout = value
        @format.send(:sync_font_ivars_from_state)
      end

      def outline
        @format.instance_variable_get(:@font_state).outline
      end

      def outline=(value)
        @format.instance_variable_get(:@font_state).outline = value
        @format.send(:sync_font_ivars_from_state)
      end

      def shadow
        @format.instance_variable_get(:@font_state).shadow
      end

      def shadow=(value)
        @format.instance_variable_get(:@font_state).shadow = value
        @format.send(:sync_font_ivars_from_state)
      end

      def script
        @format.instance_variable_get(:@font_state).script
      end

      def script=(value)
        @format.instance_variable_get(:@font_state).script = value
        @format.send(:sync_font_ivars_from_state)
      end

      def family
        @format.instance_variable_get(:@font_state).family
      end

      def family=(value)
        @format.instance_variable_get(:@font_state).family = value
        @format.send(:sync_font_ivars_from_state)
      end

      def charset
        @format.instance_variable_get(:@font_state).charset
      end

      def charset=(value)
        @format.instance_variable_get(:@font_state).charset = value
        @format.send(:sync_font_ivars_from_state)
      end

      def scheme
        @format.instance_variable_get(:@font_state).scheme
      end

      def scheme=(value)
        @format.instance_variable_get(:@font_state).scheme = value
        @format.send(:sync_font_ivars_from_state)
      end

      def condense
        @format.instance_variable_get(:@font_state).condense
      end

      def condense=(value)
        @format.instance_variable_get(:@font_state).condense = value
        @format.send(:sync_font_ivars_from_state)
      end

      def extend
        @format.instance_variable_get(:@font_state).extend
      end

      def extend=(value)
        @format.instance_variable_get(:@font_state).extend = value
        @format.send(:sync_font_ivars_from_state)
      end

      def theme
        @format.instance_variable_get(:@font_state).theme
      end

      def theme=(value)
        @format.instance_variable_get(:@font_state).theme = value
        @format.send(:sync_font_ivars_from_state)
      end

      def hyperlink
        @format.instance_variable_get(:@font_state).hyperlink
      end

      def hyperlink=(value)
        @format.instance_variable_get(:@font_state).hyperlink = value
        @format.send(:sync_font_ivars_from_state)
      end

      def color_indexed
        @format.instance_variable_get(:@font_state).color_indexed
      end

      def color_indexed=(value)
        @format.instance_variable_get(:@font_state).color_indexed = value
        @format.send(:sync_font_ivars_from_state)
      end
    end
  end
end
