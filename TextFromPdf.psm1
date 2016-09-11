################################################################################
# TextFromPdf - A PowerShell module for extracting text from PDF.
# Copyright (C) 2016 Antony Onipko
#
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU Affero General Public License as
# published by the Free Software Foundation, either version 3 of the
# License, or (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU Affero General Public License for more details.
#
# You should have received a copy of the GNU Affero General Public License
# along with this program.  If not, see <http://www.gnu.org/licenses/>.
################################################################################

Function Global:getTextFromLocation ($page, $x, $y, $w, $h) {
    $boundingBox = New-Object iText.Kernel.Geom.Rectangle($x, $y, $w, $h) # x, y, width, height
    $filter = New-Object iText.Kernel.Pdf.Canvas.Parser.Filter.TextRegionEventFilter($boundingBox)
    $strategy = New-Object iText.Kernel.Pdf.Canvas.Parser.Listener.FilteredTextEventListener(
        #(New-Object iText.Kernel.Pdf.Canvas.Parser.Listener.LocationTextExtractionStrategy),
        (New-Object iText.Kernel.Pdf.Canvas.Parser.Listener.SimpleTextExtractionStrategy),
        $filter
    )
    return [iText.Kernel.Pdf.Canvas.Parser.PdfTextExtractor]::GetTextFromPage($page, $strategy)
}

################################################################################

Function Get-TextFromPDF {
    <#
        .SYNOPSIS
        Extracts text values from PDFs using the iText 7 library.
        
        .EXAMPLE
        Get-TextFromPDF -Path 'c:\temp\receipt01.pdf'

        .EXAMPLE
        '.\receipt01.pdf', '.\receipt02.pdf' | Get-TextFromPDF
        
        .DESCRIPTION
        By default set up to work with receipts, but the extraction rules can be
        customised by passing a different RuleSet variable. Text is cleaned up using
        the TextCleanup scriptblock.
        It might be possible to obtain better results by using a different
        TraverseHeight.
    #>
    [CmdletBinding(DefaultParameterSetName='Ratio')]
    [OutputType([PsObject])]
    Param
    (
        # Path of the PDF to process.
        [Parameter(Mandatory=$true,
                   Position=1,
                   ValueFromPipeline=$True,
                   ValueFromPipelineByPropertyName=$True)]
        [Alias('FullName')]
        [ValidateScript({Test-Path $_})]
        [string]$Path,

        # An array of rules to process. Each object should contain a Name, Expression,
        # and if required a Function. The function is passed as a new closure so has
        # access to the parent function variables including private ones: text, page,
        # x, y, w, h.
        [Parameter(Mandatory=$false,
                   Position=2)]
        [PsObject[]]$RuleSet = @(
            [pscustomobject]@{
                Name       = "Date"
                Expression = "(?i)((.{2}/){2}.{2,4})|(\d{2}.?((jan(uary)?)|(feb(ruary)?)|(mar(ch?))|(apr(il)?)|(may)|(jun(e)?)|(jul(y)?)|(aug(ust)?)|(sept?(ember)?)|(oct(ober)?)|(nov(ember)?)|(dec(ember)?)).?\d{2,4})"
                Function   = $null
            }
            [pscustomobject]@{
                Name       = "Time"
                Expression = "\d{1,2}:\d{2} ?(AM|PM)?"
                Function   = $null
            }
            [pscustomobject]@{
                Name       = "Total"
                Expression = "(?i)TOTAL|TO PAY"
                Function   = {
                    $val = [regex]::Match($text, "\d{1,2}\.\d{2}").Value
                    return $val
                }
            }
            [pscustomobject]@{
                Name       = "Vat"
                Expression = "(?i)TAX|VAT(?! No)(?! Reg)(?! ?%)"
                Function   = {
                    $rx = "\d{1,2}\.\d{2}"
                    $val = [regex]::Match($text, $rx).Value
                    # Some receipts have the vat/tax in a table, with value below.
                    while(!$val -and ($x + $TraverseWidth) -le $w) {
                        $splitText = getTextFromLocation $page $x $y $TraverseWidth $h
                        if (![string]::IsNullOrWhiteSpace($splitText)) {
                            $splitText = & $TextCleanup $splitText
                            Write-Verbose "    $x,$y :  $splitText"
                            if ([regex]::Match($splitText, "(?i)TAX|VAT(?! No)(?! Reg)(?! ?%)").Success) {
                                $taxText = getTextFromLocation $page $x ($y - $TraverseHeight) $TraverseWidth $h
                                $taxText = & $TextCleanup $taxText
                                $val = [regex]::Match($taxText, $rx).Value
                            }
                        }
                        $x += $TraverseWidth
                    }
                    $x = 0
                    return $val
                }
            }
        ),

        # A scriptblock that's used to cleanup the OCR text before running the RuleSet on it.
        # By default cleans up:
        [Parameter(Mandatory=$false)]
        [scriptblock]$TextCleanup = {
            param($text)

            # Bad OCR characters
            $text = $text -replace ",", "."
            $text = $text -replace "·", "."
            $text = $text -replace "-", "."
            $text = $text -replace "'", " "

            # Spaces and new lines
            $text = $text -replace "`n", " "
            $text = $text -replace " +", " "
            $text = $text -replace " ?/ ?", "/"
            $text = $text -replace " ?: ?", ":"
            $text = $text -replace "(\d{1,2}) ?([\.:/]) ?(\d{2})",'$1$2$3' # Spaces between decimal points

            $text = $text.Trim()

            return $text
        },

        # Used to set TraverseHeight. By default:
        # TraverseHeight = RatioHeightBase * (Page.Height / Page.Width)
        [Parameter(Mandatory=$false,
                   ParameterSetName='Ratio')]
        [float]$RatioHeightBase = 15,

        # This is the height of the rectangle that extracts text.
        [Parameter(Mandatory=$false,
                   ParameterSetName='Height')]
        [float]$TraverseHeight, # Best results with a 30

        # Width of the rectangle that extracts text. Not used by default, but can be
        # used in any of the RuleSet functions. Default is 100.
        [Parameter(Mandatory=$false)]
        [float]$TraverseWidth = 100
    )

    Process {
        
        try {
            $reader = New-Object iText.Kernel.Pdf.PdfReader $Path
            $pdf = New-Object iText.Kernel.Pdf.PdfDocument $reader
        } catch {
            Write-Error $_.Exception.Message
            return
        }

        $results = [pscustomobject]@{
            Path = $Path
        }

        $RuleSet | % {
            Add-Member -InputObject $results `
                       -NotePropertyName $_.Name `
                       -NotePropertyValue ""
        }

        for ($pageNumber = 1; $pageNumber -le $pdf.GetNumberOfPages(); $pageNumber++) {
    
            $page = $pdf.GetPage($pageNumber)
            $pageSize = $page.GetPageSize()

            if (!$TraverseHeight) {
                $TraverseHeight = $RatioHeightBase * ($pageSize.GetHeight() / $pageSize.GetWidth())
            }

            [float]$x = 0
            [float]$y = $pageSize.GetHeight() - $TraverseHeight
            [float]$w = $pageSize.GetWidth()
            [float]$h = $TraverseHeight

            for ( ; $y -ge 0 ; $y -= $TraverseHeight) {

                $text = getTextFromLocation $page $x $y $w $h

                if (![string]::IsNullOrWhiteSpace($text)) {

                    $text = & $TextCleanup $text

                    Write-Verbose "$x,$y :  $text"

                    $RuleSet | ? { !$results."$($_.Name)" } | % {
                        if (($m = [regex]::Match($text, $_.Expression)).Success) {
                            if ($_.Function) {
                                $value = & $_.Function.GetNewClosure()
                            } else {
                                $value = $m.Value
                            }
                            Write-Verbose "    $($_.Name): $($value)"
                            $results."$($_.Name)" = $value
                        }
                    }

                }

            }

        }

        if ($reader) {
            $reader.Close()
        }

        Write-Output $results

    }

}

Export-ModuleMember -Function Get-TextFromPDF

################################################################################
