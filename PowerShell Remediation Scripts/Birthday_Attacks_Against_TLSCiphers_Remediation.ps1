$WeakCipherSuites = @(
    "DES",
    "IDEA",
    "RC"
)

Foreach($WeakCipherSuite in $WeakCipherSuites){
    $CipherSuites = Get-TlsCipherSuite -Name $WeakCipherSuite

    if($CipherSuites){
        Foreach($CipherSuite in $CipherSuites){
            Disable-TlsCipherSuite -Name $($CipherSuite.Name)
        }
    }
}