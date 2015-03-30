$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$sut = (Split-Path -Leaf $MyInvocation.MyCommand.Path).Replace(".Tests.", ".")
. "$here\$sut"

Describe "Get-ResourceCollectionMemberhip" {
   
    
    Context "Machine SMS Provider " {
        $computername = 'doesnotexists'
        
        It 'Should throw an error' {
            {Get-ResourceCollectionMemberhip -Name test -ComputerName $computername } | Should throw
        }

        It 'should throw an error' {
            {Get-ResourceCollectionMemberhip -Name test -ComputerName $env:COMPUTERNAME} | should throw
        }
    }

    Context "ResourceType is Device " {

         Mock -CommandName  Get-CimInstance  -MockWith {[pscustomobject]@{CollectionId='123456';Name='TestDeviceCollection';ResourceID='98765432'}}
         Mock -CommandName Get-CimInstance -ParameterFilter {$className -eq 'SMS_ProviderLocation' } -MockWith {[pscustomobject]@{NameSpace='\\TestBox.test.com\root\sms\site_TES'} }
   
        It "should spit a PSCustomObject" {
           
           $expected = Get-ResourceCollectionMemberhip -Name test 
           $expected.CollectionId | should be '123456'
           $expected.CollectionName | should be 'TestDeviceCollection'
           Assert-MockCalled -CommandName Get-CimInstance -Exactly -Times 4 -Scope Context
        }
    }
    

    Context "ResourceType is User" {
        Mock -CommandName  Get-CimInstance  -MockWith {[pscustomobject]@{CollectionId='3492512';Name='TestUserCollection';ResourceID='98765432'}}
        Mock -CommandName Get-CimInstance -ParameterFilter {$className -eq 'SMS_ProviderLocation' } -MockWith {[pscustomobject]@{NameSpace='\\TestBox.test.com\root\sms\site_TES'} }
        
        It "should spit a PSCustomObject" {
           
           $expected = Get-ResourceCollectionMemberhip -Name test -ResourceType User
           $expected.CollectionId | should be '3492512'
           $expected.CollectionName | should be 'testusercollection'
        }
    }
}
