Imports System
Imports System.Collections.Generic
Imports System.Text
Imports System.Runtime.InteropServices

Namespace FastColoredTextBoxNS
    Module PlatformType
        Const PROCESSOR_ARCHITECTURE_INTEL As UShort = 0
        Const PROCESSOR_ARCHITECTURE_IA64 As UShort = 6
        Const PROCESSOR_ARCHITECTURE_AMD64 As UShort = 9
        Const PROCESSOR_ARCHITECTURE_UNKNOWN As UShort = &HFFFF

        <StructLayout(LayoutKind.Sequential)>
        Structure SYSTEM_INFO
            Public wProcessorArchitecture As UShort
            Public wReserved As UShort
            Public dwPageSize As UInteger
            Public lpMinimumApplicationAddress As IntPtr
            Public lpMaximumApplicationAddress As IntPtr
            Public dwActiveProcessorMask As UIntPtr
            Public dwNumberOfProcessors As UInteger
            Public dwProcessorType As UInteger
            Public dwAllocationGranularity As UInteger
            Public wProcessorLevel As UShort
            Public wProcessorRevision As UShort
        End Structure

        <DllImport("kernel32.dll")>
        Private Sub GetNativeSystemInfo(ByRef lpSystemInfo As SYSTEM_INFO)
        <DllImport("kernel32.dll")>
        Private Sub GetSystemInfo(ByRef lpSystemInfo As SYSTEM_INFO)

        Function GetOperationSystemPlatform() As Platform
            Dim sysInfo = New SYSTEM_INFO()

            If Environment.OSVersion.Version.Major > 5 OrElse (Environment.OSVersion.Version.Major = 5 AndAlso Environment.OSVersion.Version.Minor >= 1) Then
                GetNativeSystemInfo(sysInfo)
            Else
                GetSystemInfo(sysInfo)
            End If

            Select Case sysInfo.wProcessorArchitecture
                Case PROCESSOR_ARCHITECTURE_IA64, PROCESSOR_ARCHITECTURE_AMD64
                    Return Platform.X64
                Case PROCESSOR_ARCHITECTURE_INTEL
                    Return Platform.X86
                Case Else
                    Return Platform.Unknown
            End Select
        End Function
    End Module

    Public Enum Platform
        X86
        X64
        Unknown
    End Enum
End Namespace