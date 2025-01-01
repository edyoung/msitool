# msitool

A command-line utility for doing various things to Windows Installer (.MSI) files.

The initial functionality is extracting embedded files stored as binaries and unpacking them.

## Background

Windows Installer files are embedded databases that contain instructions about how to install software on Windows.
Msitool can be used to read and update these database files.

The Binary table contains embedded files, usually DLLs containing custom actions. Each entry has a name and a binary blob.

## Usage

### Extracting files

```pwsh
# Extract all files from foo.msi and put them in the current directory
msitool extract -i foo.msi 

# Extract the binaries called file1,file2,file3 and create file1.dll, file2.dll, file3.dll in destinationdir
msitool extract -i foo.msi -d destinationdir [file1 file2 file3] 
```

### Inserting files

```pwsh
# Insert sourcedir\file1.dll from sourcedir into foo.msi as a binary called file1
msitool insert -i foo.msi -d sourcedir file1.dll 
```

NB: At present, if there is not already a binary row file1 in the DLL, the insert will be skipped, so this can only be used to update a binary that is already present


