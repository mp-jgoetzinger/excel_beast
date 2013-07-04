# encoding: utf-8

require 'rubygems'
require 'bundler'
begin
  Bundler.setup(:default, :development)
rescue Bundler::BundlerError => e
  $stderr.puts e.message
  $stderr.puts "Run `bundle install` to install missing gems"
  exit e.status_code
end
require 'rake'

require 'jeweler'
Jeweler::Tasks.new do |gem|
  # gem is a Gem::Specification... see http://docs.rubygems.org/read/chapter/20 for more options
  gem.name = "excel_beast"
  gem.homepage = "http://github.com/mp-jgoetzinger/excel_beast"
  gem.license = "MIT"
  gem.summary = %Q{wrapper around writeexcel that allows really simple excel generation}
  gem.description = <<-DESC.gsub('    ', '')
    a gem that aims to provide an even simpler way to generate excel files than writeexcel.
    beast_excel is probably not what you're looking for if you want to achieve more advanced stuff
    like using excel formulas, charts etc.

    but if you have some (text) data you want to export to excel, try it.
  DESC
  gem.email = "goetzinger@mediapeers.com"
  gem.authors = ["Johannes-Kostas Goetzinger"]
  # dependencies defined in Gemfile
end
Jeweler::RubygemsDotOrgTasks.new

require 'rspec/core'
require 'rspec/core/rake_task'
RSpec::Core::RakeTask.new(:spec) do |spec|
  spec.pattern = FileList['spec/**/*_spec.rb']
end

RSpec::Core::RakeTask.new(:rcov) do |spec|
  spec.pattern = 'spec/**/*_spec.rb'
  spec.rcov = true
end

task :default => :spec

require 'yard'
YARD::Rake::YardocTask.new
