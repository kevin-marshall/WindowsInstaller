require 'win32ole'
require 'cmd'

class WindowsInstaller < Hash
  def initialize(options = {})
    options.each { |key, value| self[key] = value}
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

    msi_file = msi_file.gsub(/\//, '\\')

	cmd = "msiexec.exe"
   	cmd = "#{cmd} /quiet" if(has_key?(:mode) && (self[:mode] == :quiet))
   	cmd = "#{cmd} /passive" if(has_key?(:mode) && (self[:mode] == :passive))
	cmd = "#{cmd} /i \"#{msi_file}\""

	cmd_options = {quiet: true}
	cmd_options[:admin_user] = self[:admin_user] if(has_key?(:admin_user))
    command = CMD.new(cmd, cmd_options)
	command.execute
  end

  def uninstall_msi(msi_file)
    raise "#{msi_file} does not exist!" unless(File.exists?(msi_file))
    info = msi_properties(msi_file)
    uninstall_product_code(info['ProductCode'])
  end
  
  def uninstall_product_code(product_code)
    raise "#{product_code} is not installed" unless(product_code_installed?(product_code))

	cmd = "msiexec.exe"
   	cmd = "#{cmd} /quiet" if(has_key?(:mode) && (self[:mode] == :quiet))
   	cmd = "#{cmd} /passive" if(has_key?(:mode) && (self[:mode] == :passive))
	cmd = "#{cmd} /x #{product_code}"

	cmd_options = {quiet: true}
	cmd_options[:admin_user] = self[:admin_user] if(has_key?(:admin_user))
	command = CMD.new(cmd, cmd_options)
	command.execute
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
end