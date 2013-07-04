require File.expand_path(File.dirname(__FILE__) + './../spec_helper')

describe "ExcelBeast::Sheet" do
  before(:each) do
    @workbook = WriteExcel.new(File.join(TEST_DIR, 'a'))
    @sheet = ExcelBeast::Sheet.new(@workbook, "test")
    @worksheet = @sheet.instance_variable_get("@worksheet")
  end

  context "#write" do
    it "should receive #write_cell two times" do
      @sheet.should_receive(:write_cell).exactly(2).times
      @sheet << ['foo', 'bar']
    end

    it "should increment the line index" do
      @sheet << ["ho"]
      expect(@sheet.instance_variable_get("@line_index")).to eq(1)
    end

    it "should write a numeric value" do
      @worksheet.should_receive(:write_number).once
      @sheet << [1] 
    end

    it "should write a string" do
      @worksheet.should_receive(:write_string).once
      @sheet << ["foo"] 
    end
  end

  context "#method_missing" do
    it "should delegate #write_url to Writeexcel::Worksheet" do
      @worksheet.should_receive(:write_url).once
      @sheet.write_url(0, 0, 'ftp://www.ruby.org/')
    end
  end

end
