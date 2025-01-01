using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Text;

public static class Win32ErrorHelper
{
    [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Auto)]
    private static extern int FormatMessage(
        int dwFlags, IntPtr lpSource, int dwMessageId, int dwLanguageId,
        StringBuilder lpBuffer, int nSize, IntPtr Arguments);

    private const int FORMAT_MESSAGE_ALLOCATE_BUFFER = 0x00000100;
    private const int FORMAT_MESSAGE_IGNORE_INSERTS = 0x00000200;
    private const int FORMAT_MESSAGE_FROM_SYSTEM = 0x00001000;

    public static string GetErrorMessage(int errorCode)
    {
        StringBuilder messageBuffer = new StringBuilder(256);
        int size = FormatMessage(
            FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS,
            IntPtr.Zero, errorCode, 0, messageBuffer, messageBuffer.Capacity, IntPtr.Zero);

        if (size == 0)
        {
            throw new Win32Exception(Marshal.GetLastWin32Error());
        }

        return messageBuffer.ToString().Trim();
    }
}
