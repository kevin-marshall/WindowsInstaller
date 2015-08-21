require 'rspec'
require './lib/WindowsInstaller.rb'

describe 'WindowsInstaller' do
  test_file='../test_files/example.msi'
  winstall = WindowsInstaller.new

  it 'example.msi should not be installed' do
	expect(winstall.msi_installed?(test_file).to eq(false)
  end

  it 'should be able to install example.msi' do
    winstall.install_msi(test_file)
	expect(winstall.msi_installed?(test_file).to eq(true)
  end
  
  it 'should be able to uninstall example.msi' do
    winstall.uninstall_msi(test_file)
	expect(winstall.msi_installed?(test_file).to eq(false)
  end
end