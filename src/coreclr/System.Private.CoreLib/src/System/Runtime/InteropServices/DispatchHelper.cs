// Licensed to the .NET Foundation under one or more agreements.
// The .NET Foundation licenses this file to you under the MIT license.

using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices.ComTypes;
using System.Runtime.InteropServices.Marshalling;

namespace System.Runtime.InteropServices
{
    internal struct DispatchInfo
    {
        public bool m_bInvokeUsingInvokeMember;

        public Type GetReflectionObject()
        {
            throw null;
        }
    }

    internal struct DispatchMemberInfo
    {
        public bool IsLastParamOleVarArg;
        public bool IsCultureAware;
        public bool RequiresManagedObjCleanup;

        public MemberInfo GetMemberInfoObject()
        {
            throw null;
        }

        public string GetName()
        {
            throw null;
        }

        public ParameterInfo[] GetParameters()
        {
            throw null;
        }

        public bool IsParamInOnly(int index)
        {
            throw null;
        }

        public void CleanUpParamManaged(int iParam, object? obj)
        {
            throw null;
        }
    }

    internal static class DispatchHelper
    {
        private const VarEnum VT_TYPEMASK = (VarEnum)4095;

        private static unsafe object? MarshalParamNativeToManaged(
            DispatchMemberInfo* pDispMemberInfo,
            int iParam,
            in ComVariant pSrcVar)
        {
            throw null;
        }

        private static unsafe void MarshalParamManagedToNativeRef(
            DispatchMemberInfo* pDispMemberInfo,
            int iParam,
            object? srcObj,
            object? backUpStaticArray,
            ComVariant* pRefVar)
        {
            throw null;
        }

        private static unsafe void MarshalReturnValueManagedToNative(
            DispatchMemberInfo* pDispMemberInfo,
            object? srcObj,
            ref ComVariant pDestVar)
        {
            throw null;
        }

        [RequiresUnreferencedCode("Built-in COM marshaling is incompatible with trimming.")]
        internal static unsafe void InvokeMemberWorker(
            DispatchInfo* pDispInfo,
            DispatchMemberInfo* pDispMemberInfo,
            object target,
            InvokeFlags flags,
            int dispId,
            int lcid,
            int numParams,
            int numArgs,
            int numNamedArgs,
            ref int iSrcArg,
            int* pSrcArgNames,
            ComVariant* pSrcArgs,
            ComVariant* pVarRes,
            DISPPARAMS* pdp,
            Exception* pException)
        {
            CultureInfo? oldCultureInfo = null;
            IntPtr pSA = IntPtr.Zero;
            object?[]? cleanUpArray = null;

            try
            {
                // Allocate information used by the method.
                int NumByrefArgs = 0;

                // Allocate the array of backup byref static array objects.
                object?[] aByrefStaticArrayBackupObjHandle = new object[numArgs];

                // Allocate the array that maps method params to their indices.
                Span<int> pManagedMethodParamIndexMap = stackalloc int[numArgs];

                // Allocate the array of byref objects
                ComVariant** aByrefArgOleVariant = stackalloc ComVariant*[numArgs];

                Span<bool> argUsedFlags = stackalloc bool[numParams];
                argUsedFlags.Clear();
                Span<int> aByrefArgMngVariantIndex = stackalloc int[numArgs];
                aByrefArgMngVariantIndex.Clear();

                // Allocate the array of arguments

                // Allocate the array that will contain the converted variants in the right order.
                // If the invoke is for a PROPUT or a PROPPUTREF and we are going to call through
                // invoke member then allocate the array one bigger to allow space for the property
                // value.
                int arraySize = numParams;
                if (pDispInfo->m_bInvokeUsingInvokeMember && (flags & (InvokeFlags.DISPATCH_PROPERTYPUT | InvokeFlags.DISPATCH_PROPERTYPUTREF)) != 0)
                {
                    arraySize++;
                }

                object?[] paramArray = new object[arraySize];
                object? propVal = null;
                bool propValIsByRef = false;
                int iDestArg;
                object? ByrefStaticArrayBackupPropVal = null;

                // Convert the property set argument if the invoke is a PROPERTYPUT OR PROPERTYPUTREF.
                if ((flags & (InvokeFlags.DISPATCH_PROPERTYPUT | InvokeFlags.DISPATCH_PROPERTYPUTREF)) != 0)
                {
                    // Convert the variant.
                    ComVariant* pSrcOleVariant = RetrieveSrcVariant((ComVariant*)pdp->rgvarg);
                    propVal = MarshalParamNativeToManaged(pDispMemberInfo, numArgs, in *pSrcOleVariant);

                    // Remember if the property value is byref or not.
                    propValIsByRef = pSrcOleVariant->IsByref;

                    // If the variant is a byref static array, then remember the property value.
                    if (IsVariantByrefStaticArray(pSrcOleVariant))
                    {
                        ByrefStaticArrayBackupPropVal = propVal;
                    }
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
                            if (IsVariantByrefStaticArray(pSrcOleVariant))
                            {
                                aByrefStaticArrayBackupObjHandle[NumByrefArgs] = obj;
                            }

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
                            if (IsVariantByrefStaticArray(pSrcOleVariant))
                            {
                                aByrefStaticArrayBackupObjHandle[NumByrefArgs] = obj;
                            }

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
                    ComVariant* pSrcOleVariant = null;
                    ComVariant* pFrstVarargOleVariant = null;
                    ComVariant safeArrayVar = default;
                    bool bByrefArg = false;
                    if (iDestArg == numParams - 1 && bLastParamOleVarArg)
                    {
                        // VarArg scenario
                        bool srcArgIsSafeArray = false;
                        if (iSrcArg == numNamedArgs)
                        {
                            pSrcOleVariant = RetrieveSrcVariant(&pSrcArgs[iSrcArg]);
                            if (pSrcOleVariant->VarType is (VarEnum.VT_ARRAY | VarEnum.VT_VARIANT)
                                or VarEnum.VT_ARRAY) // see the comments in case c) above
                            {
                                // vararg case c)
                                srcArgIsSafeArray = true;
                                bByrefArg = pSrcOleVariant->IsByref;
                            }
                        }

                        if (!srcArgIsSafeArray)
                        {
                            // vararg case a), b) and d)
                            // 1. Construct a safearray
                            int lSafeArrayArg = 0;
                            bByrefArg = false;
                            pSA = SafeArrayCreateVector(VarEnum.VT_VARIANT, 0, (uint)(iSrcArg - numNamedArgs + 1));
                            if (pSA == IntPtr.Zero)
                            {
                                throw new OutOfMemoryException();
                            }

                            safeArrayVar = ComVariant.CreateRaw(VarEnum.VT_VARIANT | VarEnum.VT_ARRAY, pSA);

                            // 2. Put the remaining srcArg into the safearray
                            for (; iSrcArg >= numNamedArgs; iSrcArg--, lSafeArrayArg++)
                            {
                                pSrcOleVariant = RetrieveSrcVariant(&pSrcArgs[iSrcArg]);

                                int hr = SafeArrayPutElement(pSA, &lSafeArrayArg, pSrcOleVariant);
                                Marshal.ThrowExceptionForHR(hr);

                                // Handle the UnMarshal Scenario
                                if (lSafeArrayArg == 0)
                                    pFrstVarargOleVariant = pSrcOleVariant;

                                // If any of the VARIANTS which are put into safearray is BYREF, we need marshal back it
                                bByrefArg |= pSrcOleVariant->IsByref;
                            }
                        }

                        // 3. Adjust the pSrcOleVariant in order to marshal to the params array in managed side
                        pSrcOleVariant = &safeArrayVar;
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
                        if (pSA != IntPtr.Zero)
                            aByrefArgOleVariant[NumByrefArgs] = pFrstVarargOleVariant;
                        else
                            aByrefArgOleVariant[NumByrefArgs] = pSrcOleVariant;

                        aByrefArgMngVariantIndex[NumByrefArgs] = iDestArg;

                        // If the variant is a byref static array, then remember the objectref we
                        // converted the variant to.
                        if (IsVariantByrefStaticArray(pSrcOleVariant))
                        {
                            aByrefStaticArrayBackupObjHandle[NumByrefArgs] = obj;
                        }

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

                // Set up the binding flags to pass to reflection.
                BindingFlags bindingFlags = ConvertInvokeFlagsToBindingFlags(flags) | BindingFlags.OptionalParamBinding;

                // Do the actual invocation on the member info.
                object? retVal = null;
                if (!pDispInfo->m_bInvokeUsingInvokeMember)
                {
                    Debug.Assert(pDispMemberInfo != null);

                    if (pDispMemberInfo->IsCultureAware)
                    {
                        // If the method is culture aware, then set the specified culture on the thread.
                        oldCultureInfo = CultureInfo.CurrentUICulture;
                        CultureInfo.CurrentUICulture = new CultureInfo(lcid);
                    }

                    // If the method has custom marshalers then we will need to call
                    // the clean up method on the objects. So we need to make a copy of the
                    // ParamArray since it might be changed by reflection if any of the
                    // parameters are byref.
                    if (pDispMemberInfo->RequiresManagedObjCleanup)
                    {
                        // Allocate the clean up array.
                        cleanUpArray = new object[numParams];

                        // Copy the parameters into the clean up array.
                        for (int i = 0; i < paramArray.Length; i++)
                        {
                            cleanUpArray[i] = paramArray[i];
                        }

                        // If this invoke is for a PROPUT or PROPPUTREF, then add the property object to
                        // the end of the clean up array.
                        if ((flags & (InvokeFlags.DISPATCH_PROPERTYPUT | InvokeFlags.DISPATCH_PROPERTYPUTREF)) != 0)
                        {
                            cleanUpArray[numParams] = propVal;
                        }
                    }

                    // Retrieve the member info object and the type of the member.
                    MemberInfo memberInfo = pDispMemberInfo->GetMemberInfoObject();
                    switch (memberInfo.MemberType)
                    {
                        case MemberTypes.Field:
                        {
                            FieldInfo fieldInfo = (FieldInfo)memberInfo;
                            // Make sure this invoke is actually for a property put or get.
                            if ((flags & (InvokeFlags.DISPATCH_METHOD | InvokeFlags.DISPATCH_PROPERTYGET)) != 0)
                            {
                                // Do some more validation now that we know the type of the invocation.
                                if (numNamedArgs != 0)
                                {
                                    Marshal.ThrowExceptionForHR(HResults.DISP_E_NONAMEDARGS);
                                }
                                if (numArgs != 0)
                                {
                                    Marshal.ThrowExceptionForHR(HResults.DISP_E_BADPARAMCOUNT);
                                }

                                // Do the actual method invocation.
                                retVal = fieldInfo.GetValue(target);
                            }
                            else if ((flags & (InvokeFlags.DISPATCH_PROPERTYPUT | InvokeFlags.DISPATCH_PROPERTYPUTREF)) != 0)
                            {
                                // Do some more validation now that we know the type of the invocation.
                                if (numNamedArgs != 0)
                                {
                                    Marshal.ThrowExceptionForHR(HResults.DISP_E_NONAMEDARGS);
                                }
                                if (numArgs != 0)
                                {
                                    Marshal.ThrowExceptionForHR(HResults.DISP_E_BADPARAMCOUNT);
                                }

                                // Do the actual method invocation.
                                fieldInfo.SetValue(target, propVal, bindingFlags, OleAutBinder.Instance, CultureInfo.CurrentUICulture);
                            }
                            else
                            {
                                Marshal.ThrowExceptionForHR(HResults.DISP_E_MEMBERNOTFOUND);
                            }

                            break;
                        }

                        case MemberTypes.Property:
                        {
                            PropertyInfo propertyInfo = (PropertyInfo)memberInfo;

                            // Make sure this invoke is actually for a property put or get.
                            if ((flags & (InvokeFlags.DISPATCH_METHOD | InvokeFlags.DISPATCH_PROPERTYGET)) != 0)
                            {
                                if (!IsPropertyAccessorVisible(propertyInfo, isSetter: false))
                                {
                                    Marshal.ThrowExceptionForHR(HResults.DISP_E_MEMBERNOTFOUND);
                                }

                                // Do the actual method invocation.
                                retVal = propertyInfo.GetValue(target, bindingFlags, OleAutBinder.Instance, paramArray, CultureInfo.CurrentUICulture);
                            }
                            else if ((flags & (InvokeFlags.DISPATCH_PROPERTYPUT | InvokeFlags.DISPATCH_PROPERTYPUTREF)) != 0)
                            {
                                if (!IsPropertyAccessorVisible(propertyInfo, isSetter: true))
                                {
                                    Marshal.ThrowExceptionForHR(HResults.DISP_E_MEMBERNOTFOUND);
                                }

                                // Do the actual method invocation.
                                propertyInfo.SetValue(target, propVal, bindingFlags, OleAutBinder.Instance, paramArray, CultureInfo.CurrentUICulture);
                            }
                            else
                            {
                                Marshal.ThrowExceptionForHR(HResults.DISP_E_MEMBERNOTFOUND);
                            }

                            break;
                        }

                        case MemberTypes.Method:
                        {
                            MethodBase methodInfo = (MethodBase)memberInfo;

                            // Make sure this invoke is actually for a method. We also allow
                            // prop gets since it is harmless and it allows the user a bit
                            // more freedom.
                            if ((flags & (InvokeFlags.DISPATCH_METHOD | InvokeFlags.DISPATCH_PROPERTYGET)) == 0)
                            {
                                Marshal.ThrowExceptionForHR(HResults.DISP_E_MEMBERNOTFOUND);
                            }

                            // Do the actual method invocation.
                            retVal = methodInfo.Invoke(target, bindingFlags, OleAutBinder.Instance, paramArray, CultureInfo.CurrentUICulture);
                            break;
                        }

                        default:
                        {
                            Debug.Fail("Unexpected MemberInfo type!");
                            break;
                        }
                    }
                }
                else
                {
                    // Convert the LCID into a CultureInfo.
                    CultureInfo cultureInfo = new CultureInfo(lcid);
                    Type type = pDispInfo->GetReflectionObject();

                    string memberName = pDispMemberInfo != null ? pDispMemberInfo->GetName() : $"[DISPID={dispId}]";

                    // If there are named arguments, then set up the array of named arguments
                    // to pass to InvokeMember.
                    string[]? namedArgArray = null;
                    if (numNamedArgs > 0)
                    {
                        namedArgArray = SetUpNamedParamArray(pDispMemberInfo, pSrcArgNames, numNamedArgs);
                    }

                    // If this is a PROPUT or a PROPPUTREF then we need to add the value
                    // being set as the last argument in the argument array.
                    if ((flags & (InvokeFlags.DISPATCH_PROPERTYPUT | InvokeFlags.DISPATCH_PROPERTYPUTREF)) != 0)
                    {
                        paramArray[numParams] = propVal;
                    }

                    // Do the actual method invocation.
                    retVal = type.InvokeMember(memberName, bindingFlags, OleAutBinder.Instance, target, paramArray, null, cultureInfo, namedArgArray);
                }

                // Convert the return value and the byref arguments.
                if (propValIsByRef)
                {
                    MarshalParamManagedToNativeRef(pDispMemberInfo, numArgs, propVal, ByrefStaticArrayBackupPropVal, (ComVariant*)pdp->rgvarg);
                }

                // Convert all the ByRef arguments back.
                for (int i = 0; i < NumByrefArgs; i++)
                {
                    // Get the real parameter index for this arg.
                    int iParamIndex = pManagedMethodParamIndexMap[i];

                    if (pDispMemberInfo == null || pDispInfo->m_bInvokeUsingInvokeMember || !pDispMemberInfo->IsParamInOnly(iParamIndex))
                    {
                        object? obj = paramArray[aByrefArgMngVariantIndex[i]];
                        if (pSA != IntPtr.Zero && iParamIndex == numParams - 1)
                        {
                            // VarArg scenario
                            // Here we only unmarshal the object whose corresponding VARIANT is VarArg
                            throw null;
                        }
                        else
                        {
                            MarshalParamManagedToNativeRef(pDispMemberInfo, iParamIndex, obj, aByrefStaticArrayBackupObjHandle[i], aByrefArgOleVariant[i]);
                        }
                    }
                }

                // Convert the return CLR object to an OLE variant.
                if (pVarRes != null)
                {
                    MarshalReturnValueManagedToNative(pDispMemberInfo, retVal, ref *pVarRes);
                }

            }
            catch (Exception ex)
            {
                *pException = ex;
            }
            finally
            {
                if (pSA != IntPtr.Zero)
                {
                    SafeArrayDestroy(pSA);
                }

                // If the member info requires managed object cleanup, then do it now.
                if (cleanUpArray != null)
                {
                    for (int i = 0; i < cleanUpArray.Length; i++)
                    {
                        pDispMemberInfo->CleanUpParamManaged(i, cleanUpArray[i]);
                    }
                }

                // If the culture was changed then restore it to the old culture.
                if (oldCultureInfo != null)
                {
                    CultureInfo.CurrentCulture = oldCultureInfo;
                }
            }
        }

        private static unsafe string[] SetUpNamedParamArray(DispatchMemberInfo* pDispMemberInfo, int* pSrcArgNames, int numNamedArgs)
        {
            // Allocate the array of named parameters.
            string?[] namedParamArray = new string[numNamedArgs];
            ParameterInfo[]? paramArray = pDispMemberInfo != null ? pDispMemberInfo->GetParameters() : null;

            // Convert all the named parameters from DISPID's to string.
            for (int iSrcArg = 0, iDestArg = 0; iSrcArg < numNamedArgs; iSrcArg++, iDestArg++)
            {
                // Check to see if the DISPID is one that we can map to a parameter name.
                if (pDispMemberInfo != null && pSrcArgNames[iSrcArg] >= 0 && pSrcArgNames[iSrcArg] < paramArray?.Length)
                {
                    // The DISPID is one that we assigned, map it back to its name.

                    // If we managed to get the parameters and if the current ID maps
                    // to an entry in the array.
                    if (paramArray?.Length > pSrcArgNames[iSrcArg])
                    {
                        namedParamArray[iDestArg] = paramArray[iSrcArg].Name;
                    }
                }

                // If we haven't set the param name yet, then set it to [DISP=XXXX].
                if (namedParamArray[iDestArg] == null)
                {
                    namedParamArray[iDestArg] = $"[DISP={pSrcArgNames[iSrcArg]}]";
                }
            }

            return namedParamArray!;
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

        private static BindingFlags ConvertInvokeFlagsToBindingFlags(InvokeFlags invokeFlags)
        {
            BindingFlags bindingFlags = BindingFlags.Default;

            // Check to see if DISPATCH_CONSTRUCT is set.
            if (invokeFlags.HasFlag(InvokeFlags.DISPATCH_CONSTRUCT))
            {
                bindingFlags |= BindingFlags.CreateInstance;
            }

            // Check to see if DISPATCH_METHOD is set.
            if (invokeFlags.HasFlag(InvokeFlags.DISPATCH_METHOD))
            {
                bindingFlags |= BindingFlags.InvokeMethod;
            }

            if ((invokeFlags & (InvokeFlags.DISPATCH_PROPERTYPUT | InvokeFlags.DISPATCH_PROPERTYPUTREF)) != 0)
            {
                // We are dealing with a PROPPUT or PROPPUTREF or both.
                if (invokeFlags.HasFlag(InvokeFlags.DISPATCH_PROPERTYPUT | InvokeFlags.DISPATCH_PROPERTYPUTREF))
                {
                    bindingFlags |= BindingFlags.SetProperty;
                }
                else if (invokeFlags.HasFlag(InvokeFlags.DISPATCH_PROPERTYPUT))
                {
                    bindingFlags |= BindingFlags.PutDispProperty;
                }
                else
                {
                    bindingFlags |= BindingFlags.PutRefDispProperty;
                }
            }
            else
            {
                // We are dealing with a PROPGET.
                if (invokeFlags.HasFlag(InvokeFlags.DISPATCH_PROPERTYGET))
                {
                    bindingFlags |= BindingFlags.GetProperty;
                }
            }

            return bindingFlags;
        }

        private static unsafe bool IsVariantByrefStaticArray(ComVariant* pOle)
        {
            const ushort FADF_STATIC = 0x2;

            if (pOle->IsByref && pOle->VarType.HasFlag(VarEnum.VT_ARRAY))
            {
                SafeArray* pSafeArray = *(SafeArray**)pOle->GetRawDataRef<IntPtr>();
                if (pSafeArray != null && (pSafeArray->fFeatures & FADF_STATIC) != 0)
                {
                    return true;
                }
            }

            return false;
        }

        private static bool IsPropertyAccessorVisible(PropertyInfo propertyInfo, bool isSetter)
        {
            MethodBase? accessor = isSetter ? propertyInfo.SetMethod : propertyInfo.GetMethod;
            if (accessor is null)
            {
                return false;
            }

            // TODO: handle async stub

            // Logic from IsMemberVisibleFromCom

            // Property accessor can't be generic
            Debug.Assert(accessor.IsGenericMethod);

            // Check to see if the member has the ComVisible attribute set
            return accessor.GetCustomAttribute<ComVisibleAttribute>()?.Value ?? true;
        }

        private static extern IntPtr SafeArrayCreateVector(VarEnum vt, int lLbound, uint cElements);

        private static extern unsafe int SafeArrayPutElement(IntPtr psa, int* rgIndices, void* pv);

        private static extern int SafeArrayDestroy(IntPtr psa);
    }

    internal struct SafeArray
    {
        public ushort cDims;
        public ushort fFeatures;
        public uint cbElements;
        public uint cLocks;
        public IntPtr pvData;
    }
}
