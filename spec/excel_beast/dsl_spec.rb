require File.expand_path(File.dirname(__FILE__) + './../spec_helper')

describe "ExcelBeast::Dsl" do
  context "Worksheet" do
    before(:each) do
      @worksheet = ExcelBeast::Dsl::Worksheet.new
    end

    it "should provide a setter" do
      expect(@worksheet.respond_to?(:worksheet_name)).to be_true
    end

    it "should provide a getter" do
      expect(@worksheet.respond_to?(:get_worksheet_name)).to be_true
    end

    it "should return the used value" do
      @worksheet.worksheet_name('foo')
      expect(@worksheet.get_worksheet_name).to eq('foo')
    end
  end

  context "Column" do
    before(:each) do
      @column = ExcelBeast::Dsl::Column.new
    end

    context '#to_hash' do
      it "should return a static string" do
        @column.value('bar')
        expect(@column.to_hash[:value]).to be_a(String)
      end

      it "should return a proc" do
        @column.value do
          'bar'
        end

        expect(@column.to_hash[:value]).to be_a(Proc)
      end
    end
  end

  context "sample dsl" do
    it 'should return the proper hash' do
      expected = { 
        :worksheet_name => 'test', 
        :columns => {
          'first column' => {
            :title => 'first column', # this is redundant, but it doesn't matter
            :value => 'foo',
            :format => {
              :bold => 1
            }
          }
        }
      }

      result = ExcelBeast::Dsl.evaluate do
          worksheet_name 'test'

          column do
            title 'first column'
            value 'foo'

            format do
              bold 1
            end
          end
        end

      expect(result).to eq(expected)
    end
  end
end
