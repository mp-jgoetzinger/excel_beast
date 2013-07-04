module ExcelBeast

  # class that acts as a cache for different Writeexcel::Format(s) you may use
  #
  # for all available format options follow the following link
  # @see http://writeexcel.web.fc2.com/en/doc_en.html#FORMAT_METHODS
  class Formats

    def initialize(workbook)
      @workbook = workbook
      @default_format = @workbook.add_format
      @formats = {}
    end

    # method to get a format that matches the given options.
    # creates a new one if required.
    #
    # @example
    #   #format(:bold => true)
    #
    # @params opts [Hash]
    # @return Writeexcel::Format
    # @see http://writeexcel.web.fc2.com/en/doc_en.html#FORMAT_METHODS
    def format(opts = {})
      return opts if opts.is_a?(Writeexcel::Format)
      return @default_format if !opts || opts.empty?

      key = opts.to_s
      @formats[key] ||= build_format(opts)
    end

    # returns a default (empty) format. some methods of the `writeexcel` gem require a format to be passed.
    #
    # @return [Writeexcel::Format]
    def default_format
      @default_format
    end

    # returns a format with enabled `bold` font style.
    #
    # @return [Writeexcel::Format]
    def bold_format
      format({ :bold => 1 })
    end


    private

    def build_format(opts)
      @workbook.add_format(opts)
    end

  end
end
