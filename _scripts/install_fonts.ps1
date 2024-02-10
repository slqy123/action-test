# $fonts = (New-Object -ComObject Shell.Application).Namespace(0x14)
# dir fonts/* | %{ $fonts.CopyHere($_.fullname) }

param($file)

$signature = @'
[DllImport("gdi32.dll")]
 public static extern int AddFontResource(string lpszFilename);
'@

$type = Add-Type -MemberDefinition $signature `
    -Name FontUtils -Namespace AddFontResource `
    -Using System.Text -PassThru
   
$type::AddFontResource($file)

echo OK!
