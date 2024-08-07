// Licensed to the .NET Foundation under one or more agreements.
// The .NET Foundation licenses this file to you under the MIT license.

// Generated by Fuzzlyn v2.1 on 2024-07-26 12:56:41
// Run on X64 Windows
// Seed: 13544108888657591911-vectort,vector128,vector256,x86aes,x86avx,x86avx2,x86bmi1,x86bmi1x64,x86bmi2,x86bmi2x64,x86fma,x86lzcnt,x86lzcntx64,x86pclmulqdq,x86popcnt,x86popcntx64,x86sse,x86ssex64,x86sse2,x86sse2x64,x86sse3,x86sse41,x86sse41x64,x86sse42,x86sse42x64,x86ssse3,x86x86base
// Reduced from 290.9 KiB to 0.6 KiB in 00:02:37
// Debug: Prints 1 line(s)
// Release: Prints 0 line(s)

using System;
using System.Runtime.CompilerServices;
using System.Runtime.Intrinsics;
using System.Runtime.Intrinsics.X86;
using Xunit;

public struct S0
{
    public int F0;
}

public struct S1
{
    public S0 F5;
}

public class Runtime_105558
{
    public static S1 s_3;

    [Fact]
    public static void TestEntryPoint()
    {
        if (Sse2.IsSupported)
        {
            ShiftRightLogicalTest();
        }
        if (Avx512F.IsSupported)
        {
            Avx512FRotateRightTest();
        }
    }

    [MethodImpl(MethodImplOptions.NoInlining)]
    private static void ShiftRightLogicalTest()
    {
        var vr17 = Vector128.CreateScalar(2558356441U);
        var vr18 = Vector128.Create(0, 3113514718U, 0, 0);
        var vr19 = Sse2.ShiftRightLogical(vr17, vr18);
        if (Sse2.ConvertToUInt32(vr19) != 0)
            throw new InvalidOperationException();
    }

    [MethodImpl(MethodImplOptions.NoInlining)]
    private static void Avx512FRotateRightTest()
    {
        var vr1 = Vector128.CreateScalar(1);
        Vector128<int> vr2 = Avx512F.VL.RotateRight(vr1, 84);
        if (vr2.GetElement(0) != 4096)
            throw new InvalidOperationException();
    }
}
