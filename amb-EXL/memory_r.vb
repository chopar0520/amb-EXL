Imports System.Runtime.InteropServices

Module Memory_r
    <DllImport("KERNEL32.DLL", EntryPoint:="SetProcessWorkingSetSize",
    SetLastError:=True, CallingConvention:=CallingConvention.StdCall)>
    Function SetProcessWorkingSetSize(ByVal pProcess As IntPtr,
ByVal dwMinimumWorkingSetSize As Integer,
ByVal dwMaximumWorkingSetSize As Integer) As Boolean
    End Function


    <DllImport("KERNEL32.DLL", EntryPoint:="GetCurrentProcess",
    SetLastError:=True, CallingConvention:=CallingConvention.StdCall)>
    Function GetCurrentProcess() As IntPtr
    End Function

    Sub Release()
        '↓劇的にWorking Setメモリを解放してくれる
        Dim pHandle As IntPtr = GetCurrentProcess()
        Call SetProcessWorkingSetSize(pHandle, -1, -1)
    End Sub
End Module
