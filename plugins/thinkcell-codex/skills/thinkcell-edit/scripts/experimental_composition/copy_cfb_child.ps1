# Private helper for the staged experimental composition route.
param(
 [Parameter(Mandatory=$true)][string]$Receiver,
 [Parameter(Mandatory=$true)][string]$Donor,
 [Parameter(Mandatory=$true)][string]$DonorChild,
 [Parameter(Mandatory=$true)][string]$ReceiverChild,
 [Parameter(Mandatory=$true)][string]$ModelBytes
)
$ErrorActionPreference='Stop'
Add-Type -TypeDefinition @"
using System; using System.IO; using System.Runtime.InteropServices; using System.Runtime.InteropServices.ComTypes;
[ComImport, Guid("0000000b-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface TCStorage { void CreateStream(string n,uint m,uint r1,uint r2,out IStream s); void OpenStream(string n,IntPtr r,uint m,uint r2,out IStream s); void CreateStorage(string n,uint m,uint r1,uint r2,out TCStorage s); void OpenStorage(string n,TCStorage p,uint m,IntPtr x,uint r,out TCStorage s); void CopyTo(uint n,IntPtr i,IntPtr names,TCStorage t); void MoveElementTo(string n,TCStorage t,string nn,uint f); void Commit(uint f); }
public static class TCStorageCopy { [DllImport("ole32.dll",CharSet=CharSet.Unicode)] static extern int StgOpenStorage(string p,IntPtr q,uint m,IntPtr x,uint r,out TCStorage s); static TCStorage Open(string p,uint m){TCStorage s; Marshal.ThrowExceptionForHR(StgOpenStorage(p,IntPtr.Zero,m,IntPtr.Zero,0,out s)); return s;} public static void Copy(string receiver,string donor,string donorChild,string receiverChild,string modelPath){TCStorage a=null,b=null,src=null,dst=null;IStream stream=null;IntPtr written=IntPtr.Zero;try{a=Open(receiver,0x12);b=Open(donor,0x10);b.OpenStorage(donorChild,null,0x10,IntPtr.Zero,0,out src);a.CreateStorage(receiverChild,0x12,0,0,out dst);src.CopyTo(0,IntPtr.Zero,IntPtr.Zero,dst);dst.Commit(0);a.OpenStream("think-cellXML",IntPtr.Zero,0x12,0,out stream);byte[] data=File.ReadAllBytes(modelPath);stream.Seek(0,0,IntPtr.Zero);written=Marshal.AllocCoTaskMem(4);Marshal.WriteInt32(written,0);stream.Write(data,data.Length,written);if(Marshal.ReadInt32(written)!=data.Length)throw new IOException("Partial model write");stream.SetSize(data.LongLength);stream.Commit(0);a.Commit(0);}finally{if(written!=IntPtr.Zero)Marshal.FreeCoTaskMem(written);if(stream!=null)Marshal.FinalReleaseComObject(stream);if(dst!=null)Marshal.FinalReleaseComObject(dst);if(src!=null)Marshal.FinalReleaseComObject(src);if(b!=null)Marshal.FinalReleaseComObject(b);if(a!=null)Marshal.FinalReleaseComObject(a);}} }
"@
[TCStorageCopy]::Copy($Receiver,$Donor,$DonorChild,$ReceiverChild,$ModelBytes)
