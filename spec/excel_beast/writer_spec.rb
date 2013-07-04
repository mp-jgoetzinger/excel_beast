require File.expand_path(File.dirname(__FILE__) + './../spec_helper')

# unfortunately, the output needs to be checked manually

describe "ExcelBeast::Writer" do
  context '.generate' do
    it 'should pass without errors' do
      file_path = File.join(TEST_DIR, 'test1.xls')
      definition = {
        :worksheet_name => 'Test',
        :columns => {
          'title 1' => {
            :value => ->(row) { row[:one] },
            :width => 40
          },
          'title 2' => {
            :value => ->(row) { row[:two] },
            :format => {
              :color => 'red'
            }
          },
          'title 3' => {
            :value => ->(row) { row[:three] },
            :format => {
              :color => 'green',
              :bg_color => 'silver'
            }
          }
        }
      }
      data = []
      data << { :one => 'one', :two => 'two', :three => 3 }

      ExcelBeast::Writer.generate(file_path, data, definition)
    end

    it 'should again pass without errors' do
      file_path = File.join(TEST_DIR, 'test2.xls')

      ExcelBeast::Writer.generate(file_path, ['bla']) do
        worksheet_name 'test dsl'

        column do
          title 'col 1 title'
          width 40
          value do |row|
            row
          end

          format do
            bold   1
            italic 1
          end
        end
      end
    end
  end

  context '.open' do
    it 'should pass without errors' do
      file_path = File.join(TEST_DIR, 'test3.xls')

      ExcelBeast::Writer.open(file_path) do |writer|
        sheet_one = writer.sheet("sheet one")
        sheet_two = writer.sheet("sheet two")

        bold_format = writer.formats.bold_format

        writer.use_sheet(sheet_one)
        writer.write(['title 1', 'title 2', 'title 3'], bold_format)
        writer << ['eins', 'zwo']
        writer.write_cell(1, 2, 'drei', { :bold => 1, :color => 'green' })

        writer.use_sheet("sheet two")
        writer.set_column_width(0, 50)
        writer << ['foo', 'bar']
      end
    end
  end
end
