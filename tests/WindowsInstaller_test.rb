require 'minitest/autorun'
require_relative('../lib/WindowsInstaller.rb')
require 'win32ole'

class CMD_test < MiniTest::Unit::TestCase
  def setup
    CMD.default_options({ echo_command: false, echo_output: false, debug: false})
	@installer = WindowsInstaller.new
	@test_files = {}
	@test_files[:example] = 'files/example.msi'
  end

  def test_user_install
    assert(!@installer.msi_installed?(@test_files[:example]))
	@installer.install_msi(@test_files[:example])
    assert(@installer.msi_installed?(@test_files[:example]))
	@installer.uninstall_msi(@test_files[:example])
    assert(!@installer.msi_installed?(@test_files[:example]))
  end
end