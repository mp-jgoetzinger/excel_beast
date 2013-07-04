require 'writeexcel'

Dir.glob(File.join(File.expand_path(File.dirname(__FILE__)), "excel_beast/**/*")).each do |file|
  require file unless File.directory?(file)
end

module ExcelBeast
end
