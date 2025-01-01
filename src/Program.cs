using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using CommandLine;

public class MsiTool
{
    [Verb("extract", HelpText = "Extract binaries from an MSI")]
    public class ExtractOptions
    {
        [Option(
            'i',
            "input",
            Required = true,
            HelpText = "Input MSI file, eg c:\\foo\\bar\\something.msi"
        )]
        public required string Input { get; set; }

        [Option('d', "directory", Required = false, HelpText = "Destination directory")]
        public string Destination { get; set; } = ".";

        [Option('f', "files", Required = false, HelpText = "Files to extract, omit to get all files")]
        public IEnumerable<string> Files { get; set; } = [];
    }

    [Verb("insert", HelpText = "Insert binaries into an MSI's binary table")]
    public class InsertOptions
    {
        [Option(
            'i',
            "input",
            Required = true,
            HelpText = "Input MSI file, eg c:\\foo\\bar\\something.msi"
        )]
        public required string Input { get; set; }

        [Option('d', "directory", Required = false, HelpText = "Source directory")]
        public string Destination { get; set; } = ".";

        [Option('f', "files", Required = true, HelpText = "Files to insert")]
        public IEnumerable<string> Files { get; set; } = [];
    }


    public static int Main(string[] args)
    {
        try
        {
            return CommandLine
                .Parser.Default.ParseArguments<ExtractOptions, InsertOptions>(args)
                .MapResult((ExtractOptions opts) => Extract(opts), (InsertOptions opts) => Insert(opts), errs => 1);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
            return 1;
        }
    }

    public static int Extract(ExtractOptions opts)
    {
        var msi = opts.Input;
        var dir = opts.Destination;
        var binaries = opts.Files;
        bool getall = binaries.Count() == 0;

        Console.Error.WriteLine($"Extracting binaries from {msi} into {dir}");
        int records = 0;
        int extracted = 0;

        var dbHandle = Interop.OpenDatabase(msi, DatabaseOpenMode.ReadOnly);
        try
        {
            var viewHandle = Interop.DatabaseOpenView(dbHandle, $"SELECT Name, Data FROM Binary");

            Interop.ViewExecute(viewHandle, IntPtr.Zero);

            do
            {
                var recordHandle = Interop.ViewFetch(viewHandle);

                if (recordHandle == IntPtr.Zero)
                {
                    break;
                }

                records++;
                var name = NativeMethods.GetString(recordHandle, 1);
                var size = NativeMethods.MsiRecordGetInteger(recordHandle, 2);
                if (getall || binaries.Contains(name))
                {
                    var outFile = Path.Combine(dir, $"{name}.dll");
                    
                    var bytes = new byte[size];
                    NativeMethods.MsiRecordReadStream(recordHandle, 2, bytes, ref size);

                    File.WriteAllBytes(outFile, bytes);
                    Console.Error.WriteLine($"Binary {name} size {size}: Saved to {outFile}");
                    extracted++;
                }
                else
                {
                    Console.Error.WriteLine($"Binary {name} size {size}: Skipped");
                }                
            } while (true);
        }
        finally
        {
            NativeMethods.MsiCloseHandle(dbHandle);
        }

        Console.Error.WriteLine($"Extracted {extracted} of {records} binaries");
        return 0;
    }

    public static int Insert(InsertOptions opts)
    {
        var msi = opts.Input;
        var dir = opts.Destination;
        var binaries = opts.Files;

        Console.Error.WriteLine($"Inserting files into {msi} from {dir}");

        var dbHandle = Interop.OpenDatabase(msi, DatabaseOpenMode.Transact);
        try
        {
            foreach (var name in binaries)
            {
                var inFile = Path.Combine(dir, name);
                var streamName = Path.GetFileNameWithoutExtension(inFile);
                Console.Error.WriteLine($"Updating binary record {streamName} from {inFile}");
                if (!File.Exists(inFile))
                {
                    throw new Exception($"Can't find input file {inFile}");
                }

                var query = $"UPDATE Binary SET Data = ? WHERE Name = '{streamName}'";
                var viewHandle = Interop.DatabaseOpenView(dbHandle, query);

                // Create a new record with 1 fields containing the data
                IntPtr recordHandle = NativeMethods.MsiCreateRecord(1);
                Interop.RecordSetStream(recordHandle, 1, inFile);
                Interop.ViewExecute(viewHandle, recordHandle);

                NativeMethods.MsiCloseHandle(recordHandle);
                NativeMethods.MsiCloseHandle(viewHandle);
            }

            Interop.DatabaseCommit(dbHandle);
        }
        finally
        {
            NativeMethods.MsiCloseHandle(dbHandle);
        }        
        return 0;
    }
}
