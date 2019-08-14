This is a Component Object Model (COM) Dynamic-link library (DLL) coded in C# that can be set in Classic ASP using VBscripts "CreateObject" method and allows you to analyse files of any type using ExifTool. The file information and meta data are returned as a JSON string.

ExifTool by Phil Harvey:
https://www.sno.phy.queensu.ca/~phil/exiftool/

## INSTALLATION:

Make sure you have the lastest .NET Framework installed (tested on .NET Framework 4.7.2)
	
Open IIS, go to the applicaton pools and open the application pool being used by your 
Classic ASP application. Check the .NET CRL version
E.g: v4.0.30319
	
Navigate to the CRL folder
E.g: C:\Windows\Microsoft.NET\Framework64\v4.0.30319
	
Copy over: "ClassicASP.ExifTool.dll" and "exiftool.exe"
	
Run CMD as administrator

Change the directory to your CRL folder
E.g: cd C:\Windows\Microsoft.NET\Framework64\v4.0.30319
	
Run the following command: RegAsm ClassicASP.ExifTool.dll /tlb /codebase

## Usage

	Function file_information(ByVal full_path)
				
		Dim ExifTool : Set ExifTool = Server.CreateObject("ClassicASP.ExifTool")
			
			' If the file isn't found it returns an error string: "Error: the file... couldn't be found"
			
			file_information = ExifTool.FileInfo(full_path)
			
		set ExifTool = nothing
				
	End Function

## Example output from file_information()

  [{
    "SourceFile": "C:/inetpub/wwwroot/exiftool/image.png",
    "ExifToolVersion": 11.57,
    "FileName": "image.png",
    "Directory": "C:/inetpub/wwwroot/exiftool",
    "FileSize": "4.5 MB",
    "FileModifyDate": "2019:08:14 15:34:18+01:00",
    "FileAccessDate": "2019:08:14 15:34:06+01:00",
    "FileCreateDate": "2019:08:14 15:34:17+01:00",
    "FilePermissions": "rw-rw-rw-",
    "FileType": "PNG",
    "FileTypeExtension": "png",
    "MIMEType": "image/png",
    "ImageWidth": 1600,
    "ImageHeight": 1354,
    "BitDepth": 8,
    "ColorType": "RGB with Alpha",
    "Compression": "Deflate/Inflate",
    "Filter": "Adaptive",
    "Interlace": "Noninterlaced",
    "ImageSize": "1600x1354",
    "Megapixels": 2.2
  }]
