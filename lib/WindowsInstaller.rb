require 'win32ole'
require 'cmd_windows'
require 'sys/proctable'
require 'tempfile'

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
	raise "Failed to install msi_file: #{msi_file}" unless(msi_installed?(msi_file))
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
	if(product_code_installed?(product_code))
	  properties = installed_properties(product_code)
      raise "Failed to uninstall #{properties['InstalledProductName']} #{properties['VersionString']}" 
	end
  end
  
  private
  def installed_properties(product_code)
    installer = WIN32OLE.new('WindowsInstaller.Installer')
	
	hash = Hash.new
	# known product keywords found on internet.  Would be nice to generate.
	%w[Language PackageCode Transforms AssignmentType PackageName InstalledProductName VersionString RegCompany 
	   RegOwner ProductID ProductIcon InstallLocation InstallSource InstallDate Publisher LocalPackage HelpLink 
	   HelpTelephone URLInfoAbout URLUpdateInfo InstanceType].sort.each do |prop|
	   value = installer.ProductInfo(product_code, prop)
	   hash[prop] = value unless(value.nil? || value == '')
	end
	return hash
  end
 
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
  
  def msiexec(command)	
    cmd = CMD.new(command)
	cmd.execute
  end
  
  private
  def msiexec_block
    original_value = '0'
	raise ':administrative_user must be set to setup always_install_elevated priviledges' if(self[:always_installed_elevated] && self.has_key?[:administrative_user])
	register_reg_file('1') if(self[:always_install_elevated])
	yield
	register_reg_file(current_value) if(self[:always_install_elevated])
  end

  def register_reg_file(value)
	lib_directory = File.dirname(__FILE__)
	reg_file = File.read("#{lib_directory}/always_install_elevated_template.reg")
	reg_file=reg_file.gsub(/VALUE/,value)
	file = Tempfile.new('wininst')
	file.write(reg_file)
	
	cmd = CMD.new("#{ENV['SystemRoot']}\\System32\\regedit.exe /s #{file.path}", { echo_command: true })
	cmd.execute_as(self[:administrative_user])
  end
end