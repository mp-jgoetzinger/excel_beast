require File.expand_path(File.dirname(__FILE__) + './../spec_helper')

describe "ExcelBeast::Formats" do
  before(:all) do
    @workbook = WriteExcel.new(File.join(TEST_DIR, 'a'))
  end

  it "initializes without arguing" do
    expect { ExcelBeast::Formats.new(@workbook) }.not_to raise_error
  end

  context "formats" do
    before(:each) do
      @formats = ExcelBeast::Formats.new(@workbook)
    end

    it "'s bold format has the bold option" do
      format = @formats.bold_format
      expect(format.bold).to eq(700)
    end

    it "returns the same bold format object" do
      expect(@formats.bold_format).to eq(@formats.format(:bold => 1))
    end

    it "returns a new format instance if options differ" do
      format1 = @formats.format(:color => 'red')
      format2 = @formats.format(:color => 'green')

      expect(format1).not_to eq(format2)
    end

    it "returns the format again if a format was passed instead of options" do
      format = @formats.format(:color => 'red')

      expect(format).to eq(@formats.format(format))
    end

    it "retuns the default format if #format called with empty hash" do
      expect(@formats.format({})).to eq(@formats.default_format)
    end

    it "retuns the default format if #format called with nil" do
      expect(@formats.format(nil)).to eq(@formats.default_format)
    end
  end
end
