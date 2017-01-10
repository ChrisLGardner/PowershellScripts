Function Get-PfxThumbrint
{
    [cmdletbinding()]
    param
    (
        [parameter(Mandatory)]
        [ValidateScript({
            Test-Path -Path $_
        })]
        [string]$Path,

        [parameter(Mandatory)]
        [securestring]$Password
    )

    try 
    {
        $certificateObject = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
        $certificateObject.Import($Path, $Password, [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::DefaultKeySet);

        Write-Output $certificateObject.Thumbprint
    }
    catch 
    {
        Write-Error $_
    }

}
