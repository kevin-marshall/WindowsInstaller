require 'win32ole'
require 'cmd'
require 'sys/proctable'

class WindowsInstaller < Hash
  @@default_options = {mode: '/passive'}
  
  def initialize(options = {})
    @@default_options.each { |key, value| self[key] = value }
    options.each { |key, value| self[key] = value}

	@installer = WIN32OLE.new('WindowsInstaller.Installer')
  end
  
  def self.default_options(hash) 
	hash.each { |key,value| @@default_options[key] = value }
  end
  
  def msi_installed?(msi_file)
    info = msi_properties(msi_file)
    return product_code_installed?(info['ProductCode'])
  end

  def product_installed?(product_name)
    return false if(product_name.empty?)
		
    @installer.Products.each do |product_code|
      name = @installer.ProductInfo(product_code, "ProductName")
      return product_code_installed?(product_code) if (product_name == name)
	end
		
    return false
  end
  
  def product_code_installed?(product_code)
	@installer.Products.each { |installed_product_code| return true if (product_code == installed_product_code) }
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
  
  def installed_properties(product_code)	
	hash = Hash.new
	# known product keywords found on internet.  Would be nice to generate.
	%w[Language PackageCode Transforms AssignmentType PackageName InstalledProductName VersionString RegCompany 
	   RegOwner ProductID ProductIcon InstallLocation InstallSource InstallDate Publisher LocalPackage HelpLink 
	   HelpTelephone URLInfoAbout URLUpdateInfo InstanceType].sort.each do |prop|
	   value = @installer.ProductInfo(product_code, prop)
	   hash[prop] = value unless(value.nil? || value == '')
	end
	return hash
  end
 
  private
  def msi_properties(msi_file)
    raise "#{msi_file} does not exist!" unless(File.exists?(msi_file))

    properties = {}
	  
    sql_query = "SELECT * FROM `Property`"

	db = @installer.OpenDatabase(msi_file, 0)
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

	return properties
  end
  
  def msiexec(cmd)
    cmd_options = { echo_command: false, echo_output: false} unless(self[:debug])
	if(self.has_key?(:administrative_user))
	  msiexec_admin(cmd, cmd_options)
	else
	  command = CMD.new(cmd, cmd_options)
	  command.execute
	end
  end

  def msiexec_admin(cmd, options)
    cmd = "runas /noprofile /savecred /user:#{self[:administrative_user]} \"#{cmd}\""
	command = CMD.new(cmd, options)
	wait_on_spawned(cmd) { cmd.execute }
  end
  
  def wait_on_spawned_process(cmd)
	pre_execute = Sys::ProcTable.ps
	
	pre_pids = []
	pre_execute.each { |ps| pre_pids << ps.pid }

    yield	

	exe = cmd[:command].match(/\\(?<path>.+\.exe)/i).named_captures['path']
	exe = File.basename(exe)
	#puts "Exe: #{exe}"
	
	msiexe_pid = 0
 	post_execute = Sys::ProcTable.ps
	post_execute.each do |ps| 
	  msiexe_pid = ps.pid if((ps.name.downcase == exe.downcase) && pre_pids.index(ps.pid).nil?)
	end

	if(msiexe_pid != 0)
	  loop do
	    s = Sys::ProcTable.ps(msiexe_pid)
		break if(s.nil?)
		sleep(1)
	  end
	end  
  end 
end