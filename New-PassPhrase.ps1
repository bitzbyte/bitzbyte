<#

.SYNOPSIS
Creates a new passphrase.

.EXAMPLE

.\New-PassPhrase.ps1

Generates a new passphrase that is at least 22 characters long.

#>

[CmdletBinding()]
param(
    [int]
    $Length = 22,

    [switch]
    $Uppercase,

    [switch]
    $Numbers,

    [switch]
    $Symbols,

    [switch]
    $Complex
)


function Get-WordnikWord
{
    $Word = Invoke-WebRequest -Uri http://www.wordnik.com/randoml -UseBasicParsing

    $Result = $Word.Content | Select-String -Pattern '<h1.*id="headword"[^>]*>(?<word>[^<]*)'

    Write-Output $Result.Matches.Groups[1].Value.Trim()
}


function New-ComplexCharacter
{
    [CmdletBinding()]
    param(
        [switch] $Number,
        [switch] $Symbol
    )


    $NumberSet = '0123456789'
    $SymbolSet = ' !"#$%&''()*+,-./:;<=>?@[\]^_`{|}~'

    $Set = New-Object -TypeName System.Text.StringBuilder(43)

    if ($Number) { $Set.Append($NumberSet) | Out-Null }
    if ($Symbol) { $Set.Append($SymbolSet) | Out-Null }

    Write-Output ($Set.Chars((Get-Random -Minimum 0 -Maximum ($set.Length - 1))))
}


$PassBuilder = New-Object -TypeName System.Text.StringBuilder($Length + 10)

while ($PassBuilder.Length -lt $Length)
{
    # Add a space between words
    if ($PassBuilder.Length -gt 0)
    {
        $null = $PassBuilder.Append(' ')
    }

    $Word = Get-WordnikWord
    Write-Verbose -Message "'$Word'"

    $null = $PassBuilder.Append($Word)
}


Write-Verbose -Message "Base string: $($PassBuilder.ToString())"


$ComplexChars = $PassBuilder.Length / 5

for ($i = 0; $i -lt $ComplexChars; $i++)
{
    # The character to complexify
    $Char = Get-Random -Minimum 1 -Maximum $PassBuilder.Length
    $Rand = Get-Random -Minimum 0 -Maximum 2

    # Replace a character with a symbol if 'Rand' is odd
    if ($Rand -band 1)
    {
        $c = $PassBuilder.Chars($Char)
        $NewChar = New-ComplexCharacter -Number -Symbol

        $null = $PassBuilder.Replace($c, $NewChar, $Char, 1)
        Write-Verbose -Message "Replace '$c' with '$NewChar'"
    }
    # If the character is a lowercase letter and 'Rand' is even
    elseif ($PassBuilder.Chars($Char) -cmatch '[a-z]' -and -not ($Rand -band 1))
    {
        $c = $PassBuilder.Chars($Char)
        $NewChar = [Char]::ToUpper($c)

        $null = $PassBuilder.Replace($c, $NewChar, $Char, 1)
        Write-Verbose -Message "Replace '$c' with '$NewChar'"
    }
    # If all else fails de-increment 'i' and try again
    else
    {
        $i--
    }
}

$PassBuilder.ToString()
