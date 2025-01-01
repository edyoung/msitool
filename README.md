# msitool

A command-line utility for doing various things to Windows Installer (.MSI) files.

The initial functionality is extracting embedded files stored as binaries and unpacking them.

## Background

Windows Installer files are embedded databases that contain instructions about how to install software on Windows.
Msitool can be used to read and update these database files.

## Usage

### Extracting files
```powershell 
msitool extract -i foo.msi # Extract all files from foo.msi and put them in the current directory

msitool extract -i foo.msi -d destinationdir [file1 file2 file3] # Extract the binaries called file1,file2,file3 and create file1.dll, file2.dll, file3.dll in destinationdir
```

### Inserting files
```powershell
msitool insert -i foo.msi -d sourcedir file1.dll # Insert sourcedir\file1.dll from sourcedir into foo.msi as a binary called file1
```

At present, if there is not already a binary row file1 in the DLL, the insert will be skipped


