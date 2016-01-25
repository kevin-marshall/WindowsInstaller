require 'minitest/autorun'
require_relative('../lib/WindowsInstaller.rb')
require 'win32ole'

class WindowsInstaller_test < MiniTest::Unit::TestCase
  def setup
    Execute.default_options({ echo_command: false, echo_output: false, debug: false})
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
	assert(!properties.nil?)
	
	#properties.each { |prop_name, value| puts "#{prop_name}: #{value}" }
	assert(properties['InstalledProductName'] == :example.to_s)
	assert(properties['VersionString'] == '1.0.0.0')
	assert(properties.key?('ProductCode'))
	assert(properties.key?('UpgradeCode'))
	assert(properties.key?('InstallLocation'))
  end
  
  def test_uninstall_by_property_name
	@installer.install_msi(@test_files[:example])
	assert(@installer.msi_installed?(@test_files[:example]))
	properties = @installer.installation_properties(:example.to_s)
	
	@installer.uninstall_product(properties['InstalledProductName'])
	assert(!@installer.product_installed?(@test_files[:example]))
	assert(!@installer.msi_installed?(@test_files[:example]))
  end
  
  def test_installed_products
    products = @installer.installed_products
	products.each { |product_code| assert(@installer.product_code_installed?(product_code)) }
	assert(products.size > 0)
  end
  def test_product_codes
    product_codes = @installer.product_codes('{9bc81854-26e2-4b1e-8068-48a293b1b507}')
	assert(product_codes.size == 0, 'Upgrade code does not exist, therefore there should be no assocated product codes')

	@installer.install_msi(@test_files[:example])
	msi_properites = @installer.msi_properties(@test_files[:example])
	product_codes = @installer.product_codes(msi_properites['UpgradeCode'])
	assert(product_codes.size == 1, 'There should be one product installed with the given UpgradeCode')
    assert(product_codes[0] == msi_properites['ProductCode'], 'Installed product code should equal msi product code')
	@installer.uninstall_msi(@test_files[:example])
  end

end