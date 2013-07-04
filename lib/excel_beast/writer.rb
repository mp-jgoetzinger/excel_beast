module ExcelBeast

  # a wrapper class around WriteExcel (writeexcel gem) to generate Excel files
  class Writer
    attr_reader :workbook, :file_path, :formats

    # creates a new ExcelBeast::Writer instance
    #
    # @param file_path [String] the file path the excel should be written to
    # @return [ExcelBeast::Writer]
    def initialize(file_path)
      # the path to the final excel file
      @file_path          = file_path

      # the excel workbook this instance is writing to
      # using the default font 'Arial Unicode MS' because this font properly displays
      # unicode characters on Mac OS using Microsoft Excel
      @workbook           = WriteExcel.new(file_path, :font => 'Arial Unicode MS')

      # an array of all the sheets we're working with
      @sheets             = {}

      # pointer to the current sheet data is written to
      @current_sheet      = nil

      # a Formats object holding all the formats we ever requested;
      # acts as a format cache that allows us to reuse formats with
      # the same options
      @formats            = Formats.new(@workbook)
    end

    # this method is the raison d'être for this gem.
    # provides a simplified way to write excel files just by passing data
    # and some column definitions.
    #
    # column definitions can be passed as a hash or in a dsl-ish block.
    #
    # @example example using definition hash
    #   file_path = File.join(TEST_DIR, 'test1.xls')
    #   data = []
    #   data << { :one => 'one', :two => 'two', :three => 3 }
    #   definition = {
    #     :worksheet_name => 'Test',
    #     :columns => {
    #       'title 1' => {
    #         :value => ->(row) { row[:one] },
    #         :width => 40
    #       },
    #       'title 2' => {
    #         :value => ->(row) { row[:two] },
    #         :format => {
    #           :color => 'red'
    #         }
    #       },
    #       'title 3' => {
    #         :value => ->(row) { row[:three] },
    #         :format => {
    #           :color => 'green',
    #           :bg_color => 'silver'
    #         }
    #       }
    #     }
    #   }
    #   ExcelBeast::Writer.generate(file_path, data, definition)
    #
    # @example example using DSL syntax
    #   file_path = File.join(TEST_DIR, 'test2.xls')
    #
    #   ExcelBeast::Writer.generate(file_path, ['bla']) do
    #     worksheet_name 'test dsl'
    #     column do
    #       title 'col 1 title'
    #       width 40
    #       value do |row|
    #         row
    #       end
    #       format do
    #         bold   1
    #         italic 1
    #       end
    #     end
    #   end
    #
    # @param file_path [String] the file path the excel should be written to
    # @param data [Array] the data to be written. can be anything, passed to the Proc you define for each column
    # @param definition [Hash]
    # @option definition [String] :worksheet_name name of the excel worksheet
    # @option definition [Hash] :columns Hash of Hashes; the key is used as the column title; each :column hash can use the following options
    #   * :width (Fixnum) column width 
    #   * :value (Proc, Object) can be anything; if it's a Proc, then it is called with data for each row
    #   * :format (Hash) hash containing format options, for a full list of available options see http://writeexcel.web.fc2.com/en/doc_en.html#FORMAT_METHODS
    # @param block [#to_proc] use instead of definition hash, see examples
    # @return [String] file path of the written excel file
    # @see http://writeexcel.web.fc2.com/en/doc_en.html#FORMAT_METHODS
    def self.generate(file_path, data, definition = nil, &block)
      instance = ExcelBeast::Writer.new(file_path)

      if block_given?
        definition = Dsl.evaluate(&block)
      end

      # initialize main worksheet
      instance.sheet(definition[:worksheet_name] || "Worksheet1")

      # write header
      titles = definition[:columns].keys
      instance.write(titles, instance.formats.bold_format)

      # set column width if given
      definition[:columns].each_with_index do |(key, column), index|
        next unless width = column[:width]

        instance.set_column_width(index, width)
      end

      # process data
      data.each_with_index do |row, row_index|
        # starting with line 1, as the first line is the header
        row_index += 1

        definition[:columns].each_with_index do |(key, column), column_index|
          # execute with row data if column[:value] is a Proc
          value = column[:value].is_a?(Proc) ? column[:value].call(row) : column[:value]

          # write cell
          instance.write_cell(row_index, column_index, value, column[:format])
        end
      end

      # write & close the file
      instance.close

      # return file path
      instance.file_path
    end

    # using this method you've chosen to go for the most sophisticated way of creating an Excel file
    # that this gem provides.
    #
    # on the other hand, it gives you more power (for example having more than one worksheet) while still trying
    # to make your life easier.
    #
    # @example
    #   file_path = File.join(TEST_DIR, 'test3.xls')
    #
    #   ExcelBeast::Writer.open(file_path) do |writer|
    #     sheet_one = writer.sheet("sheet one")
    #     sheet_two = writer.sheet("sheet two")
    #
    #     bold_format = writer.formats.bold_format
    #
    #     writer.use_sheet(sheet_one)
    #     writer.write(['title 1', 'title 2', 'title 3'], bold_format)
    #     writer << ['eins', 'zwo']
    #     writer.write_cell(1, 2, 'drei', { :bold => 1, :color => 'green' })
    #
    #     writer.use_sheet("sheet two")
    #     writer.set_column_width(0, 50)
    #     writer << ['foo', 'bar']
    #
    #     # you can also call original Writeexcel::Worksheet methods
    #     sheet_one.write_url(0, 0, 'ftp://www.ruby.org/')
    #
    #     # use the format caching
    #     italic_format = writer.format(:italic => true)
    #   end
    #
    # @param file_path [String] path to where the excel file should be written to
    def self.open(file_path)
      instance = ExcelBeast::Writer.new(file_path)

      yield(instance)

      instance.close
      instance.file_path
    end

    # creates a new sheet with the given name if it doesn't exist yet
    # sets the current sheet to be used for writing to this (new) sheet
    #
    # @example
    #   writer.sheet("test sheet")
    #
    # @param name [String] name of the worksheet
    # @return [ExcelBeast::Sheet]
    def sheet(name)
      @sheets[name] ||= Sheet.new(@workbook, name)
      @current_sheet = @sheets[name]
    end

    # sets the current sheet to be used for writing to the specified sheet.
    # accepts ExcelBeast::Sheet objects or a String with the sheet name.
    # creates the sheet with the specified name if it doesn't exist yet (see #sheet)
    #
    # @param name_or_sheet [ExcelBeast::Sheet, String]
    # @return [ExcelBeast::Sheet]
    def use_sheet(name_or_sheet)
      @current_sheet = name_or_sheet.is_a?(Sheet) ? name_or_sheet : sheet(name_or_sheet)
    end

    # writes one row into the excel file.
    # you may pass format options
    #
    # @param columns [Array<Object>] array of column data for this row
    # @param format_opts [Hash]
    # @see http://writeexcel.web.fc2.com/en/doc_en.html#FORMAT_METHODS
    def write(columns, format_opts = nil)
      @current_sheet.write(columns, format(format_opts))
    end
    alias_method :<<, :write

    # writes one cell of the excel file.
    # you have to pass row and column indices.
    # you may pass format options
    #
    # @param row [Fixnum] row index
    # @param column [Fixnum] column index
    # @param value [Object] what gets written to the cell
    # @param format_opts [Hash]
    # @see http://writeexcel.web.fc2.com/en/doc_en.html#FORMAT_METHODS
    def write_cell(row, column, value, format_opts = nil)
      @current_sheet.write_cell(row, column, value, format(format_opts))
    end

    # sets a column width in the current worksheet to the specified value
    #
    # @example
    #   set_column_width(0, 20)
    #
    # @param column [Fixnum] column index
    # @param value [Fixnum] width
    def set_column_width(column, value)
      @current_sheet.set_column(column, column, value)
    end

    # get format (create new or from cache) with specified options
    #
    # @param opts [Hash]
    # @return [Writeexcel::Format]
    def format(opts = {})
      formats.format(opts)
    end

    # writes excel & closes the file
    def close
      @workbook.close
    end
  end
end
