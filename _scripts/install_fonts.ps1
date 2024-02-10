$fonts = (New-Object -ComObject Shell.Application).Namespace(0x14)
dir fonts/* | %{ $fonts.CopyHere($_.fullname) }
