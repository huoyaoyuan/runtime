// Licensed to the .NET Foundation under one or more agreements.
// The .NET Foundation licenses this file to you under the MIT license.

using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices.ComTypes;
using System.Runtime.InteropServices.Marshalling;

namespace System.Runtime.InteropServices
{
    internal struct DispatchInfo
    {
        public bool m_bInvokeUsingInvokeMember;
    }

    internal struct DispatchMemberInfo
    {
        public bool IsLastParamOleVarArg;
    }

    internal static class DispatchHelper
    {
        private const ushort DISPATCH_PROPERTYPUT = 0x4;
        private const ushort DISPATCH_PROPERTYPUTREF = 0x8;
        private const VarEnum VT_TYPEMASK = (VarEnum)4095;

        private static unsafe object? MarshalParamNativeToManaged(
            DispatchMemberInfo* pDispMemberInfo,
            int iParam,
            in ComVariant pSrcVar)
        {
            throw null;
        }

        internal static unsafe void InvokeMemberWorker(
            DispatchInfo* pDispInfo,
            DispatchMemberInfo* pDispMemberInfo,
            ushort flags,
            int numParams,
            int numArgs,
            int numNamedArgs,
            ref int NumByrefArgs,
            ref int iSrcArg,
            int* pSrcArgNames,
            ComVariant* pSrcArgs,
            int* pManagedMethodParamIndexMap,
            ComVariant** aByrefArgOleVariant,
            DISPPARAMS* pdp,
            Exception* pException)
        {
            object? pSA = null; // SafeArrayPtrHolder

            try
            {
                Span<bool> argUsedFlags = stackalloc bool[numParams];
                argUsedFlags.Clear();
                Span<int> aByrefArgMngVariantIndex = stackalloc int[numArgs];
                aByrefArgMngVariantIndex.Clear();

                // Retrieve information required for the invoke call.
                _ = OleAutBinder.Instance;

                // Allocate the array of arguments

                // Allocate the array that will contain the converted variants in the right order.
                // If the invoke is for a PROPUT or a PROPPUTREF and we are going to call through
                // invoke member then allocate the array one bigger to allow space for the property
                // value.
                int arraySize = numParams;
                if (pDispInfo->m_bInvokeUsingInvokeMember && (flags & (DISPATCH_PROPERTYPUT | DISPATCH_PROPERTYPUTREF)) != 0)
                {
                    arraySize++;
                }

                object?[] paramArray = new object[arraySize];
                object? propVal = null;
                bool propValIsByRef;
                int iDestArg;

                // Convert the property set argument if the invoke is a PROPERTYPUT OR PROPERTYPUTREF.
                if ((flags & (DISPATCH_PROPERTYPUT | DISPATCH_PROPERTYPUTREF)) != 0)
                {
                    // Convert the variant.
                    ComVariant* pSrcOleVariant = RetrieveSrcVariant((ComVariant*)pdp->rgvarg);
                    propVal = MarshalParamNativeToManaged(pDispMemberInfo, numArgs, in *pSrcOleVariant);

                    // Remember if the property value is byref or not.
                    propValIsByRef = pSrcOleVariant->IsByref;

                    // If the variant is a byref static array, then remember the property value.
                    //if (IsVariantByrefStaticArray(pSrcOleVariant))
                    //    SetObjectReference(&pObjs->ByrefStaticArrayBackupPropVal, pObjs->PropVal);
                }

                // Convert the named arguments.
                if (!pDispInfo->m_bInvokeUsingInvokeMember)
                {
                    for (iSrcArg = 0; iSrcArg < numNamedArgs; iSrcArg++)
                    {
                        // Determine the destination index.
                        iDestArg = pSrcArgNames[iSrcArg];

                        // Check for duplicate param DISPID's.
                        if (argUsedFlags[iDestArg])
                        {
                            Marshal.ThrowExceptionForHR(HResults.DISP_E_PARAMNOTFOUND);
                        }

                        // Convert the variant.
                        ComVariant* pSrcOleVariant = RetrieveSrcVariant(&pSrcArgs[iSrcArg]);
                        object? obj = MarshalParamNativeToManaged(pDispMemberInfo, iDestArg, in *pSrcOleVariant);
                        paramArray[iDestArg] = obj;

                        // If the argument is byref then add it to the array of byref arguments.
                        if (pSrcOleVariant->IsByref)
                        {
                            // Remember what arg this really is.
                            pManagedMethodParamIndexMap[NumByrefArgs] = iDestArg;

                            aByrefArgOleVariant[NumByrefArgs] = pSrcOleVariant;
                            aByrefArgMngVariantIndex[NumByrefArgs] = iDestArg;

                            // If the variant is a byref static array, then remember the objectref we
                            // converted the variant to.
                            //if (IsVariantByrefStaticArray(pSrcOleVariant))
                            //    aByrefStaticArrayBackupObjHandle[NumByrefArgs] = pAppDomain->CreateHandle(pObjs->TmpObj);

                            NumByrefArgs++;
                        }

                        // Mark the slot the argument is in as occupied.
                        argUsedFlags[iDestArg] = true;
                    }
                }
                else
                {
                    for (iSrcArg = 0, iDestArg = 0; iSrcArg < numNamedArgs; iSrcArg++, iDestArg++)
                    {
                        // Check for duplicate param DISPID's.
                        if (argUsedFlags[iDestArg])
                        {
                            Marshal.ThrowExceptionForHR(HResults.DISP_E_PARAMNOTFOUND);
                        }

                        // Convert the variant.
                        ComVariant* pSrcOleVariant = RetrieveSrcVariant(&pSrcArgs[iSrcArg]);
                        object? obj = MarshalParamNativeToManaged(pDispMemberInfo, iDestArg, in *pSrcOleVariant);
                        paramArray[iDestArg] = obj;

                        // If the argument is byref then add it to the array of byref arguments.
                        if (pSrcOleVariant->IsByref)
                        {
                            // Remember what arg this really is.
                            pManagedMethodParamIndexMap[NumByrefArgs] = iDestArg;

                            aByrefArgOleVariant[NumByrefArgs] = pSrcOleVariant;
                            aByrefArgMngVariantIndex[NumByrefArgs] = iDestArg;

                            // If the variant is a byref static array, then remember the objectref we
                            // converted the variant to.
                            //if (IsVariantByrefStaticArray(pSrcOleVariant))
                            //    aByrefStaticArrayBackupObjHandle[NumByrefArgs] = pAppDomain->CreateHandle(pObjs->TmpObj);

                            NumByrefArgs++;
                        }

                        // Mark the slot the argument is in as occupied.
                        argUsedFlags[iDestArg] = true;
                    }
                }

                // Fill in the positional arguments. These are copied in reverse order and we also
                // need to skip the arguments already filled in by named arguments.
                bool bLastParamOleVarArg = pDispMemberInfo != null && pDispMemberInfo->IsLastParamOleVarArg;

                // We support VarArg by aligning with the behavior of params array in C#.
                // Here are things we do for callers depends on the arguments it passes:
                // a) NumArgs == NumParams -1:
                //     We generate a SAFEARRAY with 0 elements and pass the VARIANT
                //     wrapping it to the callee
                // b) NumArgs == NumParams && the first argument is NOT safearray:
                //     Note that arguments are passed from right to left so that the first argument
                //     passed by caller should be mapped to the last parameter of the callee
                //     We generate a SAFEARRAY to wrap the argument and pass the VARIANT
                //     wrapping the SAFEARRAY to the callee
                // c) NumArgs == NumParams && the first argument is safearray:
                //    We directly pass it to the callee. To compact with v2 behavior, we loose the
                //     conditions by checking if the VT of the safearray varaint is VT_ARRAY only
                // d) NumArgs > NumParams:
                //     We generate a SAFEARRAY to wrap then and pass the VARIANT wrapping
                //     the SAFEARRAY to the callee
                for (iSrcArg = numArgs - 1, iDestArg = 0;
                    iSrcArg >= numNamedArgs || (iDestArg == numParams - 1 && bLastParamOleVarArg)/* for vararg case a) */;
                    iSrcArg--, iDestArg++)
                {
                    // Skip the arguments already filled in by named args.
                    while (argUsedFlags[iDestArg])
                        iDestArg++;
                    Debug.Assert(iDestArg < numParams);

                    // Convert the variant.
                    ComVariant* pSrcOleVariant;
                    ComVariant* pFrstVarargOleVariant = null;
                    bool bByrefArg;
                    if (iDestArg == numParams - 1 && bLastParamOleVarArg)
                    {
                        // VarArg scenario
                        throw null;
                    }
                    else
                    {
                        pSrcOleVariant = RetrieveSrcVariant(&pSrcArgs[iSrcArg]);
                        bByrefArg = pSrcOleVariant->IsByref;
                    }

                    object? obj = MarshalParamNativeToManaged(pDispMemberInfo, iDestArg, in *pSrcOleVariant);
                    paramArray[iDestArg] = obj;

                    // If the argument is byref then add it to the array of byref arguments.
                    if (bByrefArg)
                    {
                        // Remember what arg this really is.
                        pManagedMethodParamIndexMap[NumByrefArgs] = iDestArg;

                        // Remember the original variant so that we can unmarshal it back
                        // Note that when pSA is set, pSrcOleVaraint is re-write so that we use the first argument
                        // of vararg instead
                        if (pSA != null)
                            aByrefArgOleVariant[NumByrefArgs] = pFrstVarargOleVariant;
                        else
                            aByrefArgOleVariant[NumByrefArgs] = pSrcOleVariant;

                        aByrefArgMngVariantIndex[NumByrefArgs] = iDestArg;

                        // If the variant is a byref static array, then remember the objectref we
                        // converted the variant to.
                        //if (IsVariantByrefStaticArray(pSrcOleVariant))
                        //    aByrefStaticArrayBackupObjHandle[NumByrefArgs] = pAppDomain->CreateHandle(pObjs->TmpObj);

                        NumByrefArgs++;
                    }
                }

                // Set the source arg back to -1 to indicate we are finished converting args.
                iSrcArg = -1;

                // Fill in all the remaining arguments with Missing.Value.
                for (; iDestArg < numParams; iDestArg++)
                {
                    paramArray[iDestArg] = Missing.Value;
                }
            }
            catch (Exception ex)
            {
                *pException = ex;
            }
            finally
            {
                // ManagedParamCleanupHolder
                // SafeArrayPtrHolder
            }
        }

        private static unsafe ComVariant* RetrieveSrcVariant(ComVariant* pDispParamsVariant)
        {
            // For VB6 compatibility reasons, if the VARIANT is a VT_BYREF | VT_VARIANT that
            // contains another VARIANT with VT_BYREF | VT_VARIANT, then we need to extract the
            // inner VARIANT and use it instead of the outer one. Note that if the inner VARIANT
            // is VT_BYREF | VT_VARIANT | VT_ARRAY, it will pass the below test too.
            if (pDispParamsVariant->VarType == (VarEnum.VT_BYREF | VarEnum.VT_VARIANT))
            {
                ComVariant* pByrefVariant = (ComVariant*)pDispParamsVariant->GetRawDataRef<IntPtr>();
                if ((pByrefVariant->VarType & (VT_TYPEMASK & VarEnum.VT_BYREF)) == (VarEnum.VT_VARIANT | VarEnum.VT_BYREF))
                {
                    return pByrefVariant;
                }
            }

            return pDispParamsVariant;
        }
    }
}
