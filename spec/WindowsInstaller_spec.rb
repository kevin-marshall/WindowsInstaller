require 'rspec'
require './lib/WindowsInstaller.rb'

ADMINISTRATIVE_USER='username@domain'
#ADMINISTRATIVE_USER='kmarshall@musco'

describe 'WindowsInstaller' do
  test_file='test_files/example.msi'
    
  WindowsInstaller.default_options({administrative_user: ADMINISTRATIVE_USER}) unless(ADMINISTRATIVE_USER == 'username@domain')
  
  describe 'an interactive install' do
    winstall = WindowsInstaller.new({debug: true})

    it 'example.msi should not be installed' do
	  expect(winstall.msi_installed?(test_file)).to eq(false)
    end

    it 'should be able to install example.msi' do
      winstall.install_msi(test_file)
	  expect(winstall.msi_installed?(test_file)).to eq(true)
    end
  
    it 'should be able to uninstall example.msi' do
      winstall.uninstall_msi(test_file)
	  expect(winstall.msi_installed?(test_file)).to eq(false)
    end
  end

  if(ADMINISTRATIVE_USER != 'username@domain')
    describe 'an automated install' do
      winstall = WindowsInstaller.new({mode: :quiet, debug: true})

      it 'example.msi should not be installed' do
	    expect(winstall.msi_installed?(test_file)).to eq(false)
      end

      it 'should be able to install example.msi' do
        winstall.install_msi(test_file)
	    expect(winstall.msi_installed?(test_file)).to eq(true)
      end
  
      it 'should be able to uninstall example.msi' do
        winstall.uninstall_msi(test_file)
	    expect(winstall.msi_installed?(test_file)).to eq(false)
	  end
    end
  end
end