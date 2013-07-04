module ExcelBeast

  # wrapper class around Writeexcel::Worksheet
  # provides some simplified methods to write rows and cells
  #
  # distinguishes only between numeric values and non-numeric values
  # where Writeexcel for example considers Strings starting with `=` to be an Excel formula
  class Sheet

    # creates a new sheet instance and adds itself to the given Workbook
    #
    # @param workbook [Workbook]
    # @param name [String]
    # @return [ExcelBeast::Sheet]
    def initialize(workbook, name)
      @worksheet  = workbook.add_worksheet(name)
      @line_index = 0
    end

    # writes one row into the excel file.
    #
    # @param columns [Array<Object>] array of column data for this row
    # @param format [Writeexcel::Format]
    def write(columns, format = nil)
      columns.size.times do |i|
        write_cell(@line_index, i, columns[i], format)
      end

      @line_index += 1
    end
    alias_method :<<, :write

    # writes one cell of the excel file.
    # you have to pass row and column indices.
    #
    # @param row [Fixnum] row index
    # @param column [Fixnum] column index
    # @param value [Object] what gets written to the cell
    # @param format [Writeexcel::Format]
    def write_cell(row, column, value, format = nil)
      if value.is_a?(Numeric)
        @worksheet.write_number(row, column, value, format)
      else
        @worksheet.write_string(row, column, value, format)
      end

      # return the value
      value
    end

    # delegate undefined methods to the Writeexcel::Worksheet instance
    # (if available)
    def method_missing(meth, *args, &block)
      if @worksheet.respond_to?(meth)
        @worksheet.send(meth, *args)
      else
        super
      end
    end
    
  end
end
