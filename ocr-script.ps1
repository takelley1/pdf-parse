#Extracts images from a PDF file 
#the images are converted to .jpg
function Extract-PDFImages($pdfPath,$imgFolder,$imgPrefix){
	if (!(Test-Path $imgFolder)){
		New-Item $imgFolder -ItemType Dir | Out-Null
	}
	$root="$imgFolder\$imgPrefix"
	& 'C:\misc\PDFTools\bin32\pdfimages.exe' "-j" "$pdfPath" "$root" 
	
}

Extract-PDFImages "c:\My.pdf" "c:\users\Administrator\Desktop\test" "img"






Import-Module tesseractlib.psm1
$ocr = Get-TessTextFromImage -Path "C:\Temp\test.jpg"
$ocr.Confidence
$ocr.Text






#
# Title:     tesseractlib.psm1
# Author:    Jourdan Templeton
# Email:     hello@jourdant.me
# Modified:  04/01/2015 08:30PM NZDT
#

Add-Type -AssemblyName "System.Drawing"
Add-Type -Path "$PSscriptroot\Lib\Tesseract.dll"
$tesseract = New-Object Tesseract.TesseractEngine((Get-Item "$psscriptroot\Lib\tessdata").FullName, "eng", [Tesseract.EngineMode]::Default, $null)

<#
.SYNOPSIS

This cmdlet loads either a file path or image and returns the text contained with the confidence.
.DESCRIPTION

This cmdlet loads either a file path or image and returns the text contained with the confidence.
You can pipe in either System.Drawing.Image file or a child-item object.
.PARAMETER Image

The image file already loaded into memory.
.PARAMETER FullName

The path to the image to be processed.
.EXAMPLE

$image = New-Object System.Drawing.Bitmap("c:\test.jpg")
Get-TessTextFromImage -Image $image
.EXAMPLE

New-Object System.Drawing.Bitmap("C:\test.jpg") | Get-TessTextFromImage
.EXAMPLE

$image = New-Object System.Drawing.Bitmap("c:\test.jpg")
Get-TessTextFromImage -Image $image
#>
Function Get-TessTextFromImage {
    Param(
    [Parameter(Mandatory=$true, ValueFromPipeline=$true, ParameterSetName="ImageObject")][System.Drawing.Image]$Image,
    [Parameter(Mandatory=$true, ValueFromPipeline=$true, ParameterSetName="FilePath")]    [Alias("FullName")][String]$Path
    )
	Process {
		#load image if path is a param
		If ($PsCmdlet.ParameterSetName -eq "FilePath") { $Image = New-Object System.Drawing.Bitmap((Get-Item $path).Fullname) } 

		#perform OCR on image
		$pix = [Tesseract.PixConverter]::ToPix($image)
		$page = $tesseract.Process($pix)
	
		#build return object
		$ret = New-Object PSObject -Property @{"Text"= $page.GetText();
										   "Confidence"= $page.GetMeanConfidence()}

		#clean up references
		$page.Dispose()
		If ($PsCmdlet.ParameterSetName -eq "FilePath") { $image.Dispose() } 
		return $ret
	}
}
