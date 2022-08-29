require_relative 'lib/WindowsInstaller'

test_file='test_files/example.msi'
winstall = WindowsInstaller.new

winstall.install_msi(test_file)
