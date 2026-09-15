param([Parameter(Mandatory=$true)][string]$StoragePath,[Parameter(Mandatory=$true)][string]$StreamBytesPath)
$ErrorActionPreference='Stop'
Add-Type -TypeDefinition @"
using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
[ComImport, Guid("0000000b-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface TCExperimentStorage {
 void CreateStream([MarshalAs(UnmanagedType.LPWStr)] string n, uint mode, uint r1, uint r2, out IStream stream);
 void OpenStream([MarshalAs(UnmanagedType.LPWStr)] string n, IntPtr reserved, uint mode, uint r2, out IStream stream);
 void CreateStorage([MarshalAs(UnmanagedType.LPWStr)] string n, uint mode, uint r1, uint r2, out TCExperimentStorage storage);
 void OpenStorage([MarshalAs(UnmanagedType.LPWStr)] string n, TCExperimentStorage priority, uint mode, IntPtr exclude, uint reserved, out TCExperimentStorage storage);
 void CopyTo(uint count, IntPtr excludeIids, IntPtr excludeNames, TCExperimentStorage target);
 void MoveElementTo([MarshalAs(UnmanagedType.LPWStr)] string n, TCExperimentStorage target, [MarshalAs(UnmanagedType.LPWStr)] string newName, uint flags);
 void Commit(uint flags);
}
public static class TCExperimentStreamWriter {
 [DllImport("ole32.dll", CharSet=CharSet.Unicode, PreserveSig=true)]
 static extern int StgOpenStorage(string name, IntPtr priority, uint mode, IntPtr exclude, uint reserved, out TCExperimentStorage storage);
 public static void Replace(string path, string bytesPath) {
  TCExperimentStorage st=null; IStream stream=null; IntPtr written=IntPtr.Zero;
  try {
   Marshal.ThrowExceptionForHR(StgOpenStorage(path,IntPtr.Zero,0x12,IntPtr.Zero,0,out st));
   st.OpenStream("think-cellXML",IntPtr.Zero,0x12,0,out stream);
   byte[] data=File.ReadAllBytes(bytesPath);
   stream.Seek(0,0,IntPtr.Zero);
   written=Marshal.AllocCoTaskMem(4); Marshal.WriteInt32(written,0);
   stream.Write(data,data.Length,written);
   if (Marshal.ReadInt32(written)!=data.Length) throw new IOException("Partial structured-storage write");
   stream.SetSize(data.LongLength); stream.Commit(0); st.Commit(0);
  } finally {
   if(written!=IntPtr.Zero) Marshal.FreeCoTaskMem(written);
   if(stream!=null) Marshal.FinalReleaseComObject(stream);
   if(st!=null) Marshal.FinalReleaseComObject(st);
  }
 }
}
"@
[TCExperimentStreamWriter]::Replace($StoragePath,$StreamBytesPath)
