require 'win32ole'
require 'cmd'
require 'sys/proctable'

class WindowsInstaller < Hash
  @@default_options = {mode: '/passive'}
  
  def initialize(options = {})
    @@default_options.each { |key, value| self[key] = value }
    options.each { |key, value| self[key] = value}
  end
  
  def self.default_options(hash) 
	hash.each { |key,value| @@default_options[key] = value }
  end
  
  def msi_installed?(msi_file)
    info = msi_properties(msi_file)
    return product_code_installed?(info['ProductCode'])
  end
  
  def product_code_installed?(product_code)
    installer = WIN32OLE.new('WindowsInstaller.Installer')
	installer.Products.each { |installed_product_code| return true if (product_code == installed_product_code) }
	return false
  end
  
  def install_msi(msi_file)
    raise "#{msi_file} does not exist!" unless(File.exists?(msi_file))

    msi_file = File.absolute_path(msi_file).gsub(/\//, '\\')

	cmd = "msiexec.exe"
	cmd = "#{cmd} #{self[:mode]}" if(has_key?(:mode))
	cmd = "#{cmd} /i #{msi_file}"

	msiexec(cmd)
  end

  def uninstall_msi(msi_file)
    raise "#{msi_file} does not exist!" unless(File.exists?(msi_file))
    info = msi_properties(msi_file)
    uninstall_product_code(info['ProductCode'])
  end
  
  def uninstall_product_code(product_code)
    raise "#{product_code} is not installed" unless(product_code_installed?(product_code))

	cmd = "msiexec.exe"
	cmd = "#{cmd} #{self[:mode]}" if(has_key?(:mode))
	cmd = "#{cmd} /x #{product_code}"
	msiexec(cmd)
  end
  
  private
  def msi_properties(msi_file)
    raise "#{msi_file} does not exist!" unless(File.exists?(msi_file))

    properties = {}
	  
    installer = WIN32OLE.new('WindowsInstaller.Installer')
    sql_query = "SELECT * FROM `Property`"

	db = installer.OpenDatabase(msi_file, 0)
	view = db.OpenView(sql_query)
	view.Execute(nil)
			
	record = view.Fetch()
	return nil if(record == nil)

	while(!record.nil?)
	  properties[record.StringData(1)] = record.StringData(2) 
	  record = view.Fetch()
	end
	db.ole_free
	db = nil
	  
	installer.ole_free
	installer = nil
	  
	return properties
  end
  
  def msiexec(cmd)
  	cmd = "runas /noprofile /savecred /user:#{self[:administrative_user]} \"#{cmd}\"" if(self.has_key?(:administrative_user))
	
    cmd_options = { echo_command: false, echo_output: false} unless(self[:debug])
    command = CMD.new(cmd, cmd_options)
    command.execute
  end
end