require './lib/WindowsInstaller.rb'

test_file='test_files/example.msi'
winstall = WindowsInstaller.new

winstall.install_msi(test_file)
