require 'minitest/autorun'
require_relative('../lib/WindowsInstaller.rb')
require 'win32ole'

class WindowsInstaller_test < MiniTest::Unit::TestCase
  def setup
    CMD.default_options({ echo_command: false, echo_output: false, debug: false})
	@installer = WindowsInstaller.new
	@test_files = {}
	@test_files[:example] = 'files/example.msi'
  end
  
  def teardown
	@test_files.each do |key,msi|
	  @installer.uninstall_msi(msi) if(@installer.msi_installed?(msi))
	end
  end

  def test_user_install
    assert(!@installer.msi_installed?(@test_files[:example]))
	@installer.install_msi(@test_files[:example])
    assert(@installer.msi_installed?(@test_files[:example]))
	@installer.uninstall_msi(@test_files[:example])
    assert(!@installer.msi_installed?(@test_files[:example]))
  end
  
  def test_product_installed
    assert(!@installer.product_installed?(:example.to_s))
	@installer.install_msi(@test_files[:example])
    assert(@installer.product_installed?(:example.to_s))
	@installer.uninstall_msi(@test_files[:example])
    assert(!@installer.product_installed?(:example.to_s))
  end
  
  def test_msi_properties
	properties = @installer.msi_properties(@test_files[:example])
	#properties.each { |prop_name, value| puts "#{prop_name}: #{value}" }
	assert(properties['ProductName'] == :example.to_s)
	assert(properties['ProductVersion'] == '1.0.0.0')
	assert(properties.key?('ProductCode'))
	assert(properties.key?('UpgradeCode'))
  end

  def test_installation_properties
	@installer.install_msi(@test_files[:example])
	properties = @installer.installation_properties(:example.to_s)
	#properties.each { |prop_name, value| puts "#{prop_name}: #{value}" }
	assert(properties['InstalledProductName'] == :example.to_s)
	assert(properties['VersionString'] == '1.0.0.0')
	assert(properties.key?('ProductCode'))
	assert(properties.key?('InstallLocation'))
  end
  
end