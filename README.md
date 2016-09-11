# TextFromPdf - PowerShell module for extracting text from PDF.

This module can be used to extract text from a PDF. Currently, it only contains a single function that traverses a PDF line-by-line and uses a RuleSet passed as a parameter to extract particular bits of information. It's set up to extract the total, vat, date, and time from receipts.

### Available Functions

| Function | Alias | Description |
|:---------| :---: |:------------|
| `Get-TextFromPdf` | | Extracts text from a PDF using a RuleSet. |

### Examples

##### Extract the total, tax, date, and time from a receipt
```powershell
Get-TextFromPDF -Path 'c:\temp\receipt01.pdf'
```

##### Use a custom RuleSet to extract information
```powershell
$ruleSet = @(
    [pscustomobject]@{
        Name       = "Total"
        Expression = "(?i)Net: ?"
        Function   = {
            return [regex]::Match($text, "\d{1,2}\.\d{2}").Value
        }
    }
)
Get-TextFromPDF -Path '.\receipt01.pdf' -RuleSet $ruleSet
```

### License
<a rel="license" href="https://www.gnu.org/licenses/agpl-3.0.en.html"><img alt="AGPL v3 License" style="border-width:0" src="https://www.gnu.org/graphics/agplv3-88x31.png" /> AGPL v3</a>
