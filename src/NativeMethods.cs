using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Text;

// Helper methods to wrap MSI API calls.
public static class Interop
{
    public static IntPtr OpenDatabase(string databasePath, DatabaseOpenMode persist)
    {
        int result = NativeMethods.MsiOpenDatabase(databasePath, (IntPtr)(persist), out IntPtr handle);
        if (result != 0)
        {
            throw new Exception($"OpenDatabase: MsiError {result}");
        }
        return handle;
    }

    public static IntPtr DatabaseOpenView(IntPtr dbHandle, string query)
    {
        int result = NativeMethods.MsiDatabaseOpenView(dbHandle, query, out IntPtr viewHandle);
        if (result != 0)
        {
            throw new Exception($"DatabaseOpenView: MsiError {result}");
        }
        return viewHandle;
    }

    public static void ViewExecute(IntPtr viewHandle, IntPtr recordHandle)
    {
        int result = NativeMethods.MsiViewExecute(viewHandle, recordHandle);
        if (result != 0)
        {
            throw new Exception($"ViewExecute: MsiError {result}");
        }
    }

    public static IntPtr ViewFetch(IntPtr viewHandle)
    {
        int result = NativeMethods.MsiViewFetch(viewHandle, out IntPtr recordHandle);
        if (result == 259)
        {
            // no more data
            return IntPtr.Zero;
        }
        if (result != 0)
        {
            throw new Exception($"ViewExecute: MsiError {result}");
        }
        return recordHandle;
    }

    public static void RecordSetStream(IntPtr record, int field, string filePath)
    {
        int result = NativeMethods.MsiRecordSetStream(record, field, filePath);
        if (result != 0)
        {
            throw new Exception($"SetStream: MsiError {result}");
        }
    }

    public static void DatabaseCommit(IntPtr dbHandle)
    {
        int result = NativeMethods.MsiDatabaseCommit(dbHandle);
        if (result != 0)
        {
            throw new Exception($"Commit: MsiError {result}");
        }
    }
}


// Direct wrappers for Windows Installer MSI API calls.
public static class NativeMethods
{
    /// <summary>
    /// PInvoke of MsiRecordReadStreamW.
    /// </summary>
    /// <param name="record">MSI Record handle.</param>
    /// <param name="field">Index of field to read stream from.</param>
    /// <param name="dataBuf">Data buffer to recieve stream value.</param>
    /// <param name="dataBufSize">Size of data buffer.</param>
    /// <returns>Error code.</returns>
    [DllImport("msi.dll", EntryPoint = "MsiRecordReadStream", CharSet = CharSet.Unicode, ExactSpelling = true)]
    public static extern int MsiRecordReadStream(IntPtr record, int field, byte[] dataBuf, ref int dataBufSize);

    /// <summary>
    /// PInvoke of MsiCloseHandle.
    /// </summary>
    /// <param name="database">Handle to a database.</param>
    /// <returns>Error code.</returns>
    [DllImport("msi.dll", EntryPoint = "MsiCloseHandle", CharSet = CharSet.Unicode, ExactSpelling = true)]
    public static extern int MsiCloseHandle(IntPtr database);

    /// <summary>
    /// PInvoke of MsiDatabaseOpenViewW.
    /// </summary>
    /// <param name="database">Handle to a database.</param>
    /// <param name="query">SQL query.</param>
    /// <param name="view">View handle.</param>
    /// <returns>Error code.</returns>
    [DllImport("msi.dll", EntryPoint = "MsiDatabaseOpenViewW", CharSet = CharSet.Unicode, ExactSpelling = true)]
    public static extern int MsiDatabaseOpenView(IntPtr database, string query, out IntPtr view);

    /// <summary>
    /// PInvoke of MsiOpenDatabaseW.
    /// </summary>
    /// <param name="databasePath">Path to database.</param>
    /// <param name="persist">Persist mode.</param>
    /// <param name="database">Handle to database.</param>
    /// <returns>Error code.</returns>
    [DllImport("msi.dll", EntryPoint = "MsiOpenDatabaseW", CharSet = CharSet.Unicode, ExactSpelling = true)]
    public static extern int MsiOpenDatabase(string databasePath, IntPtr persist, out IntPtr database);


    /// <summary>
    /// PInvoke of MsiRecordGetInteger.
    /// </summary>
    /// <param name="record">MSI Record handle.</param>
    /// <param name="field">Index of field to retrieve integer from.</param>
    /// <returns>Integer value.</returns>
    [DllImport("msi.dll", EntryPoint = "MsiRecordGetInteger", CharSet = CharSet.Unicode, ExactSpelling = true)]
    public static extern int MsiRecordGetInteger(IntPtr record, int field);

    /// <summary>
    /// PInvoke of MsiViewExecute.
    /// </summary>
    /// <param name="view">Handle of view to execute.</param>
    /// <param name="record">Handle to a record that supplies the parameters for the view.</param>
    /// <returns>Error code.</returns>
    [DllImport("msi.dll", EntryPoint = "MsiViewExecute", CharSet = CharSet.Unicode, ExactSpelling = true)]
    public static extern int MsiViewExecute(IntPtr view, IntPtr record);

    /// <summary>
    /// PInvoke of MsiViewFetch.
    /// </summary>
    /// <param name="view">Handle of view to fetch a row from.</param>
    /// <param name="record">Handle to receive record info.</param>
    /// <returns>Error code.</returns>
    [DllImport("msi.dll", EntryPoint = "MsiViewFetch", CharSet = CharSet.Unicode, ExactSpelling = true)]
    public static extern int MsiViewFetch(IntPtr view, out IntPtr record);

    /// <summary>
    /// PInvoke of MsiRecordGetStringW.
    /// </summary>
    /// <param name="record">MSI Record handle.</param>
    /// <param name="field">Index of field to get string value from.</param>
    /// <param name="valueBuf">Buffer to recieve value.</param>
    /// <param name="valueBufSize">Size of buffer.</param>
    /// <returns>Error code.</returns>
    [DllImport("msi.dll", EntryPoint = "MsiRecordGetStringW", CharSet = CharSet.Unicode, ExactSpelling = true)]
    public static extern int MsiRecordGetString(IntPtr record, int field, StringBuilder valueBuf, ref int valueBufSize);


    /// <summary>
    /// PInvoke of MsiRecordSetStreamW.
    /// </summary>
    /// <param name="record">MSI Record handle.</param>
    /// <param name="field">Index of field to set stream value in.</param>
    /// <param name="filePath">Path to file to set stream value to.</param>
    /// <returns>Error code.</returns>
    [DllImport("msi.dll", EntryPoint = "MsiRecordSetStreamW", CharSet = CharSet.Unicode, ExactSpelling = true)]
    internal static extern int MsiRecordSetStream(IntPtr record, int field, string filePath);

    /// <summary>
    /// PInvoke of MsiDatabaseCommit.
    /// </summary>
    /// <param name="database">Handle to a databse.</param>
    /// <returns>Error code.</returns>
    [DllImport("msi.dll", EntryPoint = "MsiDatabaseCommit", CharSet = CharSet.Unicode, ExactSpelling = true)]
    internal static extern int MsiDatabaseCommit(IntPtr database);

    /// <summary>
    /// PInvoke of MsiViewModify.
    /// </summary>
    /// <param name="view">Handle of view to modify.</param>
    /// <param name="modifyMode">Modify mode.</param>
    /// <param name="record">Handle of record.</param>
    /// <returns>Error code.</returns>
    [DllImport("msi.dll", EntryPoint = "MsiViewModify", CharSet = CharSet.Unicode, ExactSpelling = true)]
    internal static extern int MsiViewModify(IntPtr view, int modifyMode, IntPtr record);


    /// <summary>
    /// PInvoke of MsiCreateRecord
    /// </summary>
    /// <param name="parameters">Count of columns in the record.</param>
    /// <returns>Handle referencing the record.</returns>
    [DllImport("msi.dll", EntryPoint = "MsiCreateRecord", CharSet = CharSet.Unicode, ExactSpelling = true)]
    internal static extern IntPtr MsiCreateRecord(int parameters);

    /// <summary>
    /// PInvoke of MsiRecordSetStringW.
    /// </summary>
    /// <param name="record">MSI Record handle.</param>
    /// <param name="field">Index of field to set string value in.</param>
    /// <param name="value">String value.</param>
    /// <returns>Error code.</returns>
    [DllImport("msi.dll", EntryPoint = "MsiRecordSetStringW", CharSet = CharSet.Unicode, ExactSpelling = true)]
    internal static extern int MsiRecordSetString(IntPtr record, int field, string value);

    /// <summary>
    /// Gets string value at specified location.
    /// </summary>
    /// <param name="field">Index into record to get string.</param>
    /// <returns>String value</returns>
    public static string GetString(IntPtr handle, int field)
    {
        var bufferSize = 256;
        var buffer = new StringBuilder(bufferSize);
        var error = MsiRecordGetString(handle, field, buffer, ref bufferSize);
        if (234 == error)
        {
            buffer.EnsureCapacity(++bufferSize);
            error = MsiRecordGetString(handle, field, buffer, ref bufferSize);
        }

        if (0 != error)
        {
            throw new Win32Exception(error);
        }

        return (0 < buffer.Length ? buffer.ToString() : String.Empty);
    }
}

public enum DatabaseOpenMode : int
{
    /// <summary>Open a database read-only, no persistent changes.</summary>
    ReadOnly = 0,

    /// <summary>Open a database read/write in transaction mode.</summary>
    Transact = 1,

    /// <summary>Open a database direct read/write without transaction.</summary>
    Direct = 2,

    /// <summary>Create a new database, transact mode read/write.</summary>
    Create = 3,

    /// <summary>Create a new database, direct mode read/write.</summary>
    CreateDirect = 4,
}


