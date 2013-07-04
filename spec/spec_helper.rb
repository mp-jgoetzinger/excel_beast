$LOAD_PATH.unshift(File.join(File.dirname(__FILE__), '..', 'lib'))
$LOAD_PATH.unshift(File.dirname(__FILE__))
require 'rspec'
require 'excel_beast'
require 'fileutils'
require 'debugger'

# Requires supporting files with custom matchers and macros, etc,
# in ./support/ and its subdirectories.
Dir["#{File.dirname(__FILE__)}/support/**/*.rb"].each {|f| require f}

TEST_DIR = 'tmp/xls'
FileUtils.mkdir_p(TEST_DIR)

RSpec.configure do |config|
  config.mock_with :rspec
end
