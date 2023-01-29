# coding: utf-8
# frozen_string_literal: true

#
# Why would we ever use Ruby 1.8.7 when we can backport with something
# as simple as this?
#
# copied from prawn.
# modified by Hideo NAKAMURA
#
class String # :nodoc:
  def first_line
    each_line { |line| return line }
  end
  alias lines to_a unless "".respond_to?(:lines)
  unless "".respond_to?(:each_char)
    def each_char(&block) # :nodoc:
      # copied from jcode
      if block_given?
        scan(/./m, &block)
      else
        scan(/./m)
      end
    end
  end

  unless "".respond_to?(:bytesize)
    def bytesize # :nodoc:
      length
    end
  end

  unless "".respond_to?(:ord)
    def ord
      self[0]
    end
  end

  unless "".respond_to?(:ascii_only?)
    def ascii_only?
      !!(self =~ %r{[^!"#$%&'()*+,\-./:;<=>?@0-9A-Za-z_\[\\\]\{\}^` ~\0\n]})
    end
  end
end

unless File.respond_to?(:binread)
  def File.binread(file) # :nodoc:
    File.open(file, "rb:utf-8:utf-8") { |f| f.read }
  end
end

if RUBY_VERSION < "1.9"

  def ruby_18 # :nodoc:
    yield
  end

  def ruby_19 # :nodoc:
    false
  end

else

  def ruby_18 # :nodoc:
    false
  end

  def ruby_19 # :nodoc:
    yield
  end
end
