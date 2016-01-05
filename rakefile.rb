require 'dev'
require 'cmd'
require 'rbconfig'

def admin?
  `net session 2>&1`
  return ($?.to_i==0) ? true :false
end

ENV['HOME'] = ENV['USERPROFILE'] if(admin?)

task :test do
  Dir.chdir('tests') do
	Dir['*_test.rb'].each do |test|
	  cmd = CMD.new("#{RbConfig::CONFIG['bindir']}/ruby.exe #{test}", {quite: true})
	  cmd.execute
	end
  end
  
end
task :commit => [:add]

# Yard command line for realtime feed back of Readme.md modifications
# yard server --reload