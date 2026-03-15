# -*- coding: utf-8 -*-
# frozen_string_literal: true

module Writexlsx
  class Format
    class ProtectionStyle
      def initialize(format)
        @format = format
      end

      def locked
        @format.state.protection.locked
      end

      def locked=(value)
        @format.state.protection.locked = value
        @format.send(:sync_protection_ivars_from_state)
      end

      def hidden
        @format.state.protection.hidden
      end

      def hidden=(value)
        @format.state.protection.hidden = value
        @format.send(:sync_protection_ivars_from_state)
      end
    end
  end
end
