<#
.Synopsis
   Creates a new Virtual Server suitable for local lab environments.
.DESCRIPTION
   The New-LabVM cmdlet creates a new virtual server based from a pre-preared sysprep image.
.EXAMPLE
   PS C:\> New-LabVM -VMName "Server 1" -Path D:\Hyper-V -HostName localhost -Switch "Private Switch"

   This example creates a new virtual server Named "Server 1" on the local host and connects the VM to a Virtual Switch called "Private Switch".
.EXAMPLE
   PS C:\> New-LabVM -VMName "Server 2" -Path D:\Hyper-V -HostName localhost -Switch "Private Switch" -PowerOn -ConnnectVM -Verbose

   This example creates a new virtual server named "Server 2" on the localhost and connected to a virtual switch called "Private Switch".  Once created the server will be powered on and connected to.  In addition, Verbose logging is turned on.
.EXAMPLE
   PS C:\> New-LabVM -VMName "Server 3","Server 4", "Server 5" -Path D:\Hyper-V -HostName localhost -Switch "Private Switch" -PowerOn -ConnnectVM -Verbose

   This example creates three new virtual server named "Server 3", "Server 4" and "Server 5" on the localhost and connected to a virtual switch called "Private Switch".  Once created the servers will be powered on and connected to.  In addition, Verbose logging is turned on.
#>
function New-LabVM
{
    [CmdletBinding()]
    Param
    (
        # Enter Virtual Server name(s), Server names should be enclosed in quotes and multiple values should be seperated with a comma.
        [Parameter(Mandatory=$true,
                   Position=0)]
        [String[]]$VMName,
        # Enter the path for the VM to be created in.  If the path does not exist, it will be created.
        [Parameter(Mandatory=$true)]
        [String]$Path,
        # Enter the target hostname to create the VM on.  "localhost" and "." are valid for the local computer.
        [Parameter(Mandatory=$true)]
        [String]$HostName = "localhost",
        # Enter the name of a valid Virtual Switch.  The Virtual Switch name should be enclosed in quotes.
        [Parameter(Mandatory=$true)]
        [String]$Switch,
        # This will power on the newly created VM(s).
        [Switch]$PowerOn,
        # This will initiate a Console connection to the newly created VM(s).
        [Switch]$ConnnectVM
    )

    Begin
    {
        # Verifying that information entered is valid.
        # Ensure script is being run under administrative permissions:
        $wid=[System.Security.Principal.WindowsIdentity]::GetCurrent()
        $prp=new-object System.Security.Principal.WindowsPrincipal($wid)
        $adm=[System.Security.Principal.WindowsBuiltInRole]::Administrator
        $IsAdmin=$prp.IsInRole($adm)
        if (!($IsAdmin))
        {
            write-error "You must have administrative permissions to complete this process" -ErrorVariable $VMError -ErrorAction Stop
        }

        # Verify Hyper-V Host:        
        if(Get-VMHost $hostname -ErrorAction Stop -ErrorVariable $VMError)
        {       
            Write-Verbose "Hyper-V host verified." 
        }

        # Verify Virtual Switch:
        if(Get-VMSwitch -Name "$switch" -ErrorAction Stop -ErrorVariable $VMError)
        {
            Write-Verbose "Virtual Switch verified."
        }

        # Check if path contains a valid Drive letter i.e. "x:\"
        if(!(test-path $path.substring(0,3)))
        {
            write-error -Message "Can not determine valid drive.  Exiting" -ErrorAction Stop
        }
        write-verbose "Destination drive verified"
    }
    Process
    {
        foreach ($V in $VMName)
        {
            # Check if VMName is already in use 
            # If a duplicate name is found, add a incremental number until a unique name is assigned.        
            if(get-vm -ComputerName $hostname -Name $v -ErrorAction SilentlyContinue -ErrorVariable $VMError)
            {
                Write-Warning "Virtual Server $V already exists"
                Write-Warning "Finding unique server name for $V"
                $i = 1
                DO
                {
                    $i++
                }
                WHILE
                    (Get-VM -ComputerName $hostname -Name "$v$i" -ErrorAction SilentlyContinue -ErrorVariable $VMError)  
                    Write-Verbose "Changed servername $v to $v$i"
                    $v = "$v$i"   
            }            

            # Validate Path
            # Adding trailing "\" if missing.
            if(!($path.EndsWith("\")))
                {
                    $path = $path+"\"
                }
  
                # Build and validate full path if the current path exists.
                if (!(test-path $path$v))
                {
                    # Create Path
                    Write-warning "Path $path$v does not exist"
                    write-verbose "Creating $path$v"
                    New-item $Path$V -ItemType Directory | out-null -ErrorVariable $VMerror -ErrorAction SilentlyContinue
                    
                    # Verify Path
                    write-verbose "Verifying that $path$v is a valid folder" 
                    if (test-path $path$v -PathType Container)
                    {
                        write-verbose "Folder $path$v Created"
                    }
                    ELSE
                    {
                        write-error -Message "Cannot verify folder - Stopping Script" -ErrorAction Stop
                    }
                    Write-Verbose "Checking to ensure that $path$v\$v.vhdx does not exist."
                    if (!(test-path "$path$v\$v.vhdx"))
                    {
                        # Creating the VM
                        Write-verbose "Copying Sysprep'd VHDX from library"
                        copy-item 'D:\Sysprep Images\Windows Server 2012 R2 Standard - Sysprep.vhdx' "$Path$V\$v.vhdx"
                        write-verbose "Creating VM"
                        new-vm -Name $V -ComputerName $Hostname -SwitchName $switch -VHDPath "$Path$V\$v.vhdx" -Generation 2 -path $path$vm -MemoryStartupBytes 1GB | out-null
                        
                        # Verifying the VM
                        if(!(get-vm -ComputerName $hostname -Name $v -ErrorAction SilentlyContinue -ErrorVariable $VMError))
                        {
                            write-warning "Virtual Server $V cannot be verified"
                        }
                        ELSE
                        {
                            write-verbose "Virtual Server $v verified"
                        }

                        if($PowerOn)
                        {
                            # Powering on the VM
                            write-verbose "Powering on Virtual Server $V"
                            $PowerOn = start-vm -Name $V
                        }

                        if($ConnnectVM)
                        {
                            # Connecting to VM.
                            Write-verbose "Establishing Connection with Virtual Server $V"
                            vmconnect.exe $hostname $V
                        }
                    }
     
                }
                 ELSE
                 {
                    write-warning "$path$v\$v.vhdx already exists.  Skipping VM creation."
                 }   
        }
    }
    End
    {
    }
}