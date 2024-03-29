# Copyright 2013-2015 Kevin Marshall
#
#   Licensed under the Apache License, Version 2.0 (the "License");
#   you may not use this file except in compliance with the License.
#   You may obtain a copy of the License at
#
#       http://www.apache.org/licenses/LICENSE-2.0
#
#   Unless required by applicable law or agreed to in writing, software
#   distributed under the License is distributed on an "AS IS" BASIS,
#   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
#   See the License for the specific language governing permissions and
#   limitations under the License.
#
#
Gem::Specification.new do |s|
	s.name			= 'WindowsInstaller'
	s.version		= '0.1.29'
	s.summary		= 'Class wrapper for Microsoft Windows WindowsInstaller object'
	s.description	= 'Class wrapper for Microsoft Windows WindowsInstaller object'
	s.authors		= ["Kevin Marshall"]
	#s.email			= 'KCCKSMarshall@gmail.com'
	#s.rubyforge_project = 'WindowsInstaller'
	s.homepage		= 'http://rubygems.org/gems/WindowsExecute'
    s.required_ruby_version = '>= 3.0.0'
	s.files         = Dir["LICENSE","README.md","lib/**/*","example/**/*"]
	s.license       = 'Apache 2.0'
	s.require_path  = 'lib'
	s.add_dependency 'execute', '~> 0.1.79'
    s.add_development_dependency 'bundler', '~> 0'
	s.add_development_dependency 'rake', '~> 0'
	s.add_development_dependency 'rspec', '~> 0'
	s.add_development_dependency 'dev', '~> 0'
end