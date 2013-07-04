module ExcelBeast
  # set of classes that allow to pass options dsl-alike
  module Dsl

    # shared methods to that generates getters and setters
    # as well as to_hash converion
    class Base

      # writes setters and getters for the given attributes
      #
      # @param attrs [Array<Symbol>]
      def self.setup(*attrs)
        attrs.each do |attr|
          define_method(attr) do |value|
            instance_variable_set("@#{attr}", value)
          end

          define_method("get_#{attr}") do
            instance_variable_get("@#{attr}")
          end
        end
      end

      # converts the set of all instance variables to a hash.
      # if an instance variable is a Dsl::Base call #to_hash instead.
      #
      # @return [Hash]
      def to_hash
        hsh = {}
        instance_variables.each do |var|
          name = var.to_s.gsub('@','').to_sym
          obj = instance_variable_get(var)
          hsh[name] = obj.is_a?(Base) ? obj.to_hash : obj
        end

        hsh
      end
    end

    class Column < Base
      setup :title, :width

      def format(&block)
        @format = Format.new
        @format.instance_eval(&block)
      end

      # method to set a Proc or a common value for this column
      def value(value = nil, &block)
        @value = block_given? ? Proc.new(&block) : value
      end
    end

    class Format < Base
      setup :align, :bg_color, :bold, :border, :border_color, :bottom, :bottom_color,
        :center_across, :color, :fg_color, :font, :font_outline, :font_script,
        :font_shadow, :font_strikeout, :hidden, :indent, :italic, :left, :left_color,
        :locked, :num_format, :pattern, :right, :right_color, :rotation, :shrink, :size,
        :text_justlast, :text_wrap, :top, :top_color, :underline
    end

    class Worksheet < Base
      setup :worksheet_name

      def column(&block)
        column = Column.new
        column.instance_eval(&block)

        @columns ||= {}
        @columns[column.get_title] = column.to_hash
      end

      def to_hash
        {
          :worksheet_name => @worksheet_name,
          :columns => @columns
        }
      end
    end

    # method that takes a block and converts it to a definition hash
    def self.evaluate(&block)
      worksheet = Worksheet.new
      worksheet.instance_eval(&block)
      worksheet.to_hash
    end
  end
end
