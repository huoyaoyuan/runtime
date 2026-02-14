// Licensed to the .NET Foundation under one or more agreements.
// The .NET Foundation licenses this file to you under the MIT license.
//
// File: DispatchInfo.cpp
//

//
// Implementation of helpers used to expose IDispatch
// and IDispatchEx to COM.
//


#include "common.h"

#include "dispatchinfo.h"
#include "dispex.h"
#include "object.h"
#include "field.h"
#include "method.hpp"
#include "class.h"
#include "comcallablewrapper.h"
#include "threads.h"
#include "excep.h"
#include "comutilnative.h"
#include "eeconfig.h"
#include "interoputil.h"
#include "olevariant.h"
#include "commtmemberinfomap.h"
#include "dispparammarshaler.h"
#include "reflectioninvocation.h"
#include "dbginterface.h"
#include "dllimport.h"

#define EXCEPTION_INNER_PROP                            "InnerException"

// The name of the properties accessed on the managed member infos.
#define MEMBER_INFO_NAME_PROP                           "Name"

// The initial size of the DISPID to member map.
#define DISPID_TO_MEMBER_MAP_INITIAL_SIZE        37

// The names of the properties that are accessed on the managed member info's
#define MEMBERINFO_TYPE_PROP            "MemberType"

// The names of the properties that are accessed on managed ParameterInfo.
#define PARAMETERINFO_NAME_PROP         "Name"


MethodTable*      DispatchMemberInfo::s_pMemberTypes[NUM_MEMBER_TYPES]  = {NULL};
EnumMemberTypes DispatchMemberInfo::s_memberTypes[NUM_MEMBER_TYPES] = {Uninitted};
int           DispatchMemberInfo::s_iNumMemberTypesKnown            = 0;

// Helper function to convert between a DISPID and a hashkey.
inline UPTR DispID2HashKey(DISPID DispID)
{
    LIMITED_METHOD_CONTRACT;

    return DispID + 2;
}

// Typedef for string comparison functions.
typedef int (*UnicodeStringCompareFuncPtr)(const WCHAR *, const WCHAR *);

//--------------------------------------------------------------------------------
// The DispatchMemberInfo class implementation.

DispatchMemberInfo::DispatchMemberInfo(DispatchInfo *pDispInfo, DISPID DispID, SString& strName)
: m_DispID(DispID)
, m_hndMemberInfo(NULL)
, m_apParamMarshaler(NULL)
, m_pParamInOnly(NULL)
, m_pNext(NULL)
, m_strName(strName)
, m_enumType (Uninitted)
, m_iNumParams(-1)
, m_CultureAwareState(Unknown)
, m_bRequiresManagedCleanup(FALSE)
, m_bInitialized(FALSE)
, m_bNeutered(FALSE)
, m_pDispInfo(pDispInfo)
, m_bLastParamOleVarArg(FALSE)
{
    WRAPPER_NO_CONTRACT;
}

void DispatchMemberInfo::Neuter()
{
    WRAPPER_NO_CONTRACT;

    if (m_apParamMarshaler)
    {
        // Need to delete all individual members?
        // Can't calculate the exact number here.
        delete [] m_apParamMarshaler;
        m_apParamMarshaler = NULL;
    }

    if (m_pParamInOnly)
    {
        delete [] m_pParamInOnly;
        m_pParamInOnly = NULL;
    }

    //m_pNext = NULL;
    m_enumType = Uninitted;
    m_iNumParams = -1;
    m_CultureAwareState = Unknown;
    m_bNeutered = TRUE;
}

DispatchMemberInfo::~DispatchMemberInfo()
{
    CONTRACTL
    {
        NOTHROW;
        GC_NOTRIGGER;
        MODE_ANY;
    }
    CONTRACTL_END;

    // Delete the parameter marshalers and then delete the array of parameter
    // marshalers itself.
    if (m_apParamMarshaler)
    {
        EnumMemberTypes MemberType = GetMemberType();
        int NumParamMarshalers = GetNumParameters() + ((MemberType == Property) ? 2 : 1);
        for (int i = 0; i < NumParamMarshalers; i++)
        {
            if (m_apParamMarshaler[i])
                delete m_apParamMarshaler[i];
        }
        delete []m_apParamMarshaler;
    }

    if (m_pParamInOnly)
        delete [] m_pParamInOnly;

    if (m_hndMemberInfo)
        m_pDispInfo->FreeHandle(m_hndMemberInfo);

    // Clear the name of the member.
    m_strName.Clear();
}

void DispatchMemberInfo::EnsureInitialized()
{
    CONTRACTL
    {
        THROWS;
        GC_TRIGGERS;
        MODE_COOPERATIVE;
    }
    CONTRACTL_END;

    // Initialize the entry if it hasn't been initialized yet. This must be synchronized.
    if (!m_bInitialized)
    {
        DispatchInfo::LockHolder lh(m_pDispInfo);

        if (!m_bInitialized)
            Init();
    }
}

void DispatchMemberInfo::Init()
{
    CONTRACTL
    {
        THROWS;
        GC_TRIGGERS;
        MODE_COOPERATIVE;
    }
    CONTRACTL_END;

    EX_TRY
    {
        // Determine the type of the member.
        DetermineMemberType();

        // Determine the parameter count.
        DetermineParamCount();

        // Determine the culture awareness of the member.
        DetermineCultureAwareness();

        // Set up the parameter marshaler info.
        SetUpParamMarshalerInfo();

        // Mark the dispatch member info as having been initialized.
        m_bInitialized = TRUE;
    }
    EX_CATCH
    {
        // If we do throw an exception, then the status of the object
        // is in limbo - just neuter it.
        Neuter();
        RethrowTerminalExceptions();
    }
    EX_END_CATCH
}

HRESULT DispatchMemberInfo::GetIDsOfParameters(_In_reads_(NumNames) WCHAR **astrNames, int NumNames, DISPID *aDispIds, BOOL bCaseSensitive)
{
    CONTRACTL
    {
        THROWS;
        GC_TRIGGERS;
        MODE_COOPERATIVE;
        INJECT_FAULT(COMPlusThrowOM());

        // The member info must have been initialized before this is called.
        PRECONDITION(TRUE == m_bInitialized);
        PRECONDITION(CheckPointer(astrNames));
        PRECONDITION(CheckPointer(aDispIds));
    }
    CONTRACTL_END;

    int NumNamesMapped = 0;
    PTRARRAYREF ParamArray = NULL;
    int cNames = 0;

    // Initialize all the ID's to DISPID_UNKNOWN.
    for (cNames = 0; cNames < NumNames; cNames++)
        aDispIds[cNames] = DISPID_UNKNOWN;

    // Retrieve the appropriate string comparation function.
    UnicodeStringCompareFuncPtr StrCompFunc = bCaseSensitive ? u16_strcmp : SString::_wcsicmp;

    GCPROTECT_BEGIN(ParamArray)
    {
        // Retrieve the member parameters.
        ParamArray = GetParameters();

        // If we managed to retrieve an non empty array of parameters then go through it and
        // map the specified names to ID's.
        if ((ParamArray != NULL) && (ParamArray->GetNumComponents() > 0))
        {
            int NumParams = ParamArray->GetNumComponents();
            int cParams = 0;
            NewArrayHolder< NewArrayHolder<WCHAR> > astrParamNames = new NewArrayHolder<WCHAR>[NumParams];

            // Go through and retrieve the names of all the components.
            for (cParams = 0; cParams < NumParams; cParams++)
            {
                OBJECTREF ParamInfoObj = ParamArray->GetAt(cParams);
                GCPROTECT_BEGIN(ParamInfoObj)
                {
                    // Retrieve the MD to use to retrieve the name of the parameter.
                    MethodDesc *pGetParamNameMD = MemberLoader::FindPropertyMethod(ParamInfoObj->GetMethodTable(), PARAMETERINFO_NAME_PROP, PropertyGet);
                    _ASSERTE(pGetParamNameMD && "Unable to find getter method for property ParameterInfo::Name");
                    MethodDescCallSite getParamName(pGetParamNameMD, &ParamInfoObj);

                    // Retrieve the name of the parameter.
                    ARG_SLOT GetNameArgs[] =
                    {
                        ObjToArgSlot(ParamInfoObj)
                    };
                    STRINGREF MemberNameObj = getParamName.Call_RetSTRINGREF(GetNameArgs);

                        // If we got a valid name back then store that in the array of names.
                    if (MemberNameObj != NULL)
                    {
                        astrParamNames[cParams] = new WCHAR[MemberNameObj->GetStringLength() + 1];
                        wcscpy_s(astrParamNames[cParams], MemberNameObj->GetStringLength() + 1, MemberNameObj->GetBuffer());
                    }
                }
                GCPROTECT_END();
            }

            // Now go through the list of specfiied names and map then to ID's.
            for (cNames = 0; cNames < NumNames; cNames++)
            {
                for (cParams = 0; cParams < NumParams; cParams++)
                {
                    if (astrParamNames[cParams] && (StrCompFunc(astrNames[cNames], astrParamNames[cParams]) == 0))
                    {
                        aDispIds[cNames] = cParams;
                        NumNamesMapped++;
                        break;
                    }
                }
            }
        }
    }
    GCPROTECT_END();

    return (NumNamesMapped == NumNames) ? S_OK : DISP_E_UNKNOWNNAME;
}

PTRARRAYREF DispatchMemberInfo::GetParameters()
{
    CONTRACTL
    {
        THROWS;
        GC_TRIGGERS;
        MODE_COOPERATIVE;
    }
    CONTRACTL_END;

    PTRARRAYREF ParamArray = NULL;
    MethodDesc *pGetParamsMD = NULL;

    // Retrieve the method to use to retrieve the array of parameters.
    switch (GetMemberType())
    {
        case Method:
        {
            pGetParamsMD = DispatchInfo::GetMethodInfoMD(METHOD__METHOD__GET_PARAMETERS, GetMemberInfoObject()->GetTypeHandle());
            _ASSERTE(pGetParamsMD && "Unable to find method MemberBase::GetParameters");
            break;
        }

        case Property:
        {
            pGetParamsMD = DispatchInfo::GetPropertyInfoMD(METHOD__PROPERTY__GET_INDEX_PARAMETERS, GetMemberInfoObject()->GetTypeHandle());
            _ASSERTE(pGetParamsMD && "Unable to find method PropertyInfo::GetIndexParameters");
            break;
        }
    }

    // If the member has parameters then retrieve the array of parameters.
    if (pGetParamsMD != NULL)
    {
        OBJECTREF memberInfoObject = GetMemberInfoObject();
        GCPROTECT_BEGIN(memberInfoObject)
        MethodDescCallSite getParams(pGetParamsMD, &memberInfoObject);

        ARG_SLOT GetParamsArgs[] =
        {
            ObjToArgSlot(memberInfoObject)
        };

        ParamArray = (PTRARRAYREF) getParams.Call_RetOBJECTREF(GetParamsArgs);
        GCPROTECT_END();
    }

    return ParamArray;
}

OBJECTREF DispatchMemberInfo::GetMemberInfoObject()
{
    WRAPPER_NO_CONTRACT;
    return m_pDispInfo->GetHandleValue(m_hndMemberInfo);
}

void DispatchMemberInfo::MarshalParamNativeToManaged(int iParam, VARIANT *pSrcVar, OBJECTREF *pDestObj)
{
    CONTRACTL
    {
        THROWS;
        GC_TRIGGERS;
        MODE_COOPERATIVE;
        PRECONDITION(CheckPointer(pSrcVar));
        PRECONDITION(pDestObj != NULL);
        PRECONDITION(TRUE == m_bInitialized);
    }
    CONTRACTL_END;

    if (m_apParamMarshaler && m_apParamMarshaler[iParam + 1])
        m_apParamMarshaler[iParam + 1]->MarshalNativeToManaged(pSrcVar, pDestObj);
    else
        OleVariant::MarshalObjectForOleVariant(pSrcVar, pDestObj);
}

void DispatchMemberInfo::MarshalParamManagedToNativeRef(int iParam, OBJECTREF *pSrcObj, VARIANT *pRefVar)
{
    CONTRACTL
    {
        THROWS;
        GC_TRIGGERS;
        MODE_COOPERATIVE;
        PRECONDITION(CheckPointer(pRefVar));
        PRECONDITION(TRUE == m_bInitialized);
        PRECONDITION(pSrcObj != NULL);
    }
    CONTRACTL_END;

    if (m_apParamMarshaler && m_apParamMarshaler[iParam + 1])
        m_apParamMarshaler[iParam + 1]->MarshalManagedToNativeRef(pSrcObj, pRefVar);
    else
        OleVariant::MarshalOleRefVariantForObject(pSrcObj, pRefVar);
}

void DispatchMemberInfo::CleanUpParamManaged(int iParam, OBJECTREF *pObj)
{
    CONTRACTL
    {
        THROWS;
        GC_TRIGGERS;
        MODE_COOPERATIVE;
        PRECONDITION(pObj != NULL);
        PRECONDITION(TRUE == m_bInitialized);
    }
    CONTRACTL_END;

    if (m_apParamMarshaler && m_apParamMarshaler[iParam + 1])
        m_apParamMarshaler[iParam + 1]->CleanUpManaged(pObj);
}

void DispatchMemberInfo::MarshalReturnValueManagedToNative(OBJECTREF *pSrcObj, VARIANT *pDestVar)
{
    CONTRACTL
    {
        THROWS;
        GC_TRIGGERS;
        MODE_COOPERATIVE;
        PRECONDITION(CheckPointer(pDestVar));
        PRECONDITION(pSrcObj != NULL);
        PRECONDITION(TRUE == m_bInitialized);
    }
    CONTRACTL_END;

    if (m_apParamMarshaler && m_apParamMarshaler[0])
        m_apParamMarshaler[0]->MarshalManagedToNative(pSrcObj, pDestVar);
    else
        OleVariant::MarshalOleVariantForObject(pSrcObj, pDestVar);
}

ComMTMethodProps * DispatchMemberInfo::GetMemberProps(OBJECTREF MemberInfoObj, ComMTMemberInfoMap *pMemberMap)
{
    CONTRACT (ComMTMethodProps*)
    {
        THROWS;
        GC_TRIGGERS;
        MODE_COOPERATIVE;
        PRECONDITION(MemberInfoObj != NULL);
        PRECONDITION(CheckPointer(pMemberMap, NULL_OK));
        POSTCONDITION(CheckPointer(RETVAL, NULL_OK));
    }
    CONTRACT_END;

    DISPID DispId = DISPID_UNKNOWN;
    ComMTMethodProps *pMemberProps = NULL;

    // If we don't have a member map then we cannot retrieve properties for the member.
    if (!pMemberMap)
        RETURN NULL;

    // Get the member's properties.
    GCPROTECT_BEGIN(MemberInfoObj);
    {
        MethodTable *pMemberInfoClass = MemberInfoObj->GetMethodTable();
        if (CoreLibBinder::IsClass(pMemberInfoClass, CLASS__METHOD))
        {
            // Retrieve the MethodDesc from the MethodInfo.
            MethodDescCallSite getMethodHandle(METHOD__METHOD_BASE__GET_METHODDESC, &MemberInfoObj);
            ARG_SLOT GetMethodHandleArg = ObjToArgSlot(MemberInfoObj);
            MethodDesc* pMeth = (MethodDesc*) getMethodHandle.Call_RetLPVOID(&GetMethodHandleArg);
            if (pMeth)
            {
                // We don't expose runtime-async methods via IDispatch.
                if (pMeth->IsAsyncMethod())
                    RETURN NULL;

                pMemberProps = pMemberMap->GetMethodProps(pMeth->GetMemberDef(), pMeth->GetModule());
            }
        }
        else if (CoreLibBinder::IsClass(pMemberInfoClass, CLASS__RT_FIELD_INFO))
        {
            MethodDescCallSite getFieldDesc(METHOD__RTFIELD__GET_FIELDESC, &MemberInfoObj);
            ARG_SLOT arg = ObjToArgSlot(MemberInfoObj);
            FieldDesc* pFld = (FieldDesc*) getFieldDesc.Call_RetLPVOID(&arg);
            if (pFld)
                pMemberProps = pMemberMap->GetMethodProps(pFld->GetMemberDef(), pFld->GetModule());
        }
        else if (CoreLibBinder::IsClass(pMemberInfoClass, CLASS__PROPERTY))
        {
            MethodDescCallSite getToken(METHOD__PROPERTY__GET_TOKEN, &MemberInfoObj);
            ARG_SLOT arg = ObjToArgSlot(MemberInfoObj);
            mdToken propTok = (mdToken) getToken.Call_RetArgSlot(&arg);
            MethodDescCallSite getModule(METHOD__PROPERTY__GET_MODULE, &MemberInfoObj);
            ARG_SLOT arg1 = ObjToArgSlot(MemberInfoObj);
            REFLECTMODULEBASEREF module = (REFLECTMODULEBASEREF) getModule.Call_RetOBJECTREF(&arg1);
            Module* pModule = module->GetModule();
            pMemberProps = pMemberMap->GetMethodProps(propTok, pModule);
        }
    }
    GCPROTECT_END();

    RETURN pMemberProps;
}

DISPID DispatchMemberInfo::GetMemberDispId(OBJECTREF MemberInfoObj, ComMTMemberInfoMap *pMemberMap)
{
    CONTRACTL
    {
        THROWS;
        GC_TRIGGERS;
        MODE_COOPERATIVE;
        PRECONDITION(CheckPointer(pMemberMap, NULL_OK));
    }
    CONTRACTL_END;

    _ASSERT(MemberInfoObj);

    DISPID DispId = DISPID_UNKNOWN;

    // Get the member's properties.
    ComMTMethodProps *pMemberProps = GetMemberProps(MemberInfoObj, pMemberMap);

    // If we managed to get the properties of the member then extract the DISPID.
    if (pMemberProps)
        DispId = pMemberProps->dispid;

    return DispId;
}

LPWSTR DispatchMemberInfo::GetMemberName(OBJECTREF MemberInfoObj, ComMTMemberInfoMap *pMemberMap)
{
    CONTRACT (LPWSTR)
    {
        THROWS;
        GC_TRIGGERS;
        MODE_COOPERATIVE;
        INJECT_FAULT(COMPlusThrowOM());
        PRECONDITION(MemberInfoObj != NULL);
        PRECONDITION(CheckPointer(pMemberMap, NULL_OK));
        POSTCONDITION(CheckPointer(RETVAL));
    }
    CONTRACT_END;

    NewArrayHolder<WCHAR> strMemberName = NULL;
    ComMTMethodProps *pMemberProps = NULL;

    GCPROTECT_BEGIN(MemberInfoObj);
    {
        // Get the member's properties.
        pMemberProps = GetMemberProps(MemberInfoObj, pMemberMap);

        // If we managed to get the member's properties then extract the name.
        if (pMemberProps)
        {
            int MemberNameLen = (INT)u16_strlen(pMemberProps->pName);
            strMemberName = new WCHAR[MemberNameLen + 1];

            memcpy(strMemberName, pMemberProps->pName, (MemberNameLen + 1) * sizeof(WCHAR));
        }
        else
        {
            // Retrieve the Get method for the Name property.
            MethodDesc *pMD = MemberLoader::FindPropertyMethod(MemberInfoObj->GetMethodTable(), MEMBER_INFO_NAME_PROP, PropertyGet);
            _ASSERTE(pMD && "Unable to find getter method for property MemberInfo::Name");
            MethodDescCallSite propGet(pMD, &MemberInfoObj);

            // Prepare the arguments.
            ARG_SLOT Args[] =
            {
                ObjToArgSlot(MemberInfoObj)
            };

            // Retrieve the value of the Name property.
            STRINGREF strObj = propGet.Call_RetSTRINGREF(Args);
            _ASSERTE(strObj != NULL);

            // Copy the name into the buffer we will return.
            int MemberNameLen = strObj->GetStringLength();
            strMemberName = new WCHAR[strObj->GetStringLength() + 1];
            memcpy(strMemberName, strObj->GetBuffer(), MemberNameLen * sizeof(WCHAR));
            strMemberName[MemberNameLen] = 0;
        }
    }
    GCPROTECT_END();

    strMemberName.SuppressRelease();
    RETURN strMemberName;
}

void DispatchMemberInfo::DetermineMemberType()
{
    CONTRACTL
    {
        THROWS;
        GC_TRIGGERS;
        MODE_COOPERATIVE;

    // This should not be called more than once.
        PRECONDITION(m_enumType == Uninitted);
    }
    CONTRACTL_END;

    OBJECTREF MemberInfoObj = GetMemberInfoObject();

    // Check to see if the member info is of a type we have already seen.
    TypeHandle pMemberInfoClass   = MemberInfoObj->GetTypeHandle();
    for (int i = 0 ; i < s_iNumMemberTypesKnown ; i++)
    {
        if (pMemberInfoClass.GetMethodTable() == s_pMemberTypes[i])
        {
            m_enumType = s_memberTypes[i];
            return;
        }
    }

    GCPROTECT_BEGIN(MemberInfoObj);
    {
        // Retrieve the method descriptor for the type property accessor.
        MethodDesc *pMD = MemberLoader::FindPropertyMethod(MemberInfoObj->GetMethodTable(), MEMBERINFO_TYPE_PROP, PropertyGet);
        _ASSERTE(pMD && "Unable to find getter method for property MemberInfo::Type");
        MethodDescCallSite propGet(pMD, &MemberInfoObj);

        // Prepare the arguments that will be used to retrieve the value of all the properties.
        ARG_SLOT Args[] =
        {
            ObjToArgSlot(MemberInfoObj)
        };

        // Retrieve the actual type of the member info.
        m_enumType = (EnumMemberTypes)propGet.Call_RetArgSlot(Args);
    }
    GCPROTECT_END();

    if (s_iNumMemberTypesKnown < NUM_MEMBER_TYPES)
    {
        s_pMemberTypes[s_iNumMemberTypesKnown] = MemberInfoObj->GetMethodTable();
        s_memberTypes[s_iNumMemberTypesKnown++] = m_enumType;
    }
}

void DispatchMemberInfo::DetermineParamCount()
{
    CONTRACTL
    {
        THROWS;
        GC_TRIGGERS;
        MODE_COOPERATIVE;
    // This should not be called more than once.
        PRECONDITION(m_iNumParams == -1);
    }
    CONTRACTL_END;

    MethodDesc *pGetParamsMD = NULL;

    OBJECTREF MemberInfoObj = GetMemberInfoObject();
    GCPROTECT_BEGIN(MemberInfoObj);
    {
        // Retrieve the method to use to retrieve the array of parameters.
        switch (GetMemberType())
        {
            case Method:
            {
                pGetParamsMD = DispatchInfo::GetMethodInfoMD(METHOD__METHOD__GET_PARAMETERS, GetMemberInfoObject()->GetTypeHandle());
                _ASSERTE(pGetParamsMD && "Unable to find method MemberBase::GetParameters");
                break;
            }

            case Property:
            {
                pGetParamsMD = DispatchInfo::GetPropertyInfoMD(METHOD__PROPERTY__GET_INDEX_PARAMETERS, GetMemberInfoObject()->GetTypeHandle());
                _ASSERTE(pGetParamsMD && "Unable to find method PropertyInfo::GetIndexParameters");
                break;
            }
        }

        // If the member has parameters then get their count.
        if (pGetParamsMD != NULL)
        {
            MethodDescCallSite getParams(pGetParamsMD, &MemberInfoObj);

            ARG_SLOT GetParamsArgs[] =
            {
                ObjToArgSlot(GetMemberInfoObject())
            };

            PTRARRAYREF ParamArray = (PTRARRAYREF) getParams.Call_RetOBJECTREF(GetParamsArgs);
            if (ParamArray != NULL)
                m_iNumParams = ParamArray->GetNumComponents();
        }
        else
        {
            m_iNumParams = 0;
        }
    }
    GCPROTECT_END();
}

void DispatchMemberInfo::DetermineCultureAwareness()
{
    CONTRACTL
    {
        THROWS;
        GC_TRIGGERS;
        MODE_COOPERATIVE;
    // This should not be called more than once.
        PRECONDITION(m_CultureAwareState == Unknown);
    }
    CONTRACTL_END;

    // Load the LCIDConversionAttribute type.
    MethodTable * pLcIdConvAttrClass = CoreLibBinder::GetClass(CLASS__LCID_CONVERSION_TYPE);

    // Check to see if the attribute is set.
    OBJECTREF MemberInfoObj = GetMemberInfoObject();
    GCPROTECT_BEGIN(MemberInfoObj);
    {
        // Retrieve the method to use to determine if the DispIdAttribute custom attribute is set.
        MethodDesc *pGetCustomAttributesMD = DispatchInfo::GetCustomAttrProviderMD(MemberInfoObj->GetTypeHandle());
        MethodDescCallSite getCustomAttributes(pGetCustomAttributesMD, &MemberInfoObj);

        // Prepare the arguments.
        ARG_SLOT GetCustomAttributesArgs[] =
        {
            0,
            ObjToArgSlot(pLcIdConvAttrClass->GetManagedClassObject()),
            0,
        };

        // Now that we have potentially triggered a GC in the GetManagedClassObject
        // call above, it is safe to set the 'this' using our properly protected
        // MemberInfoObj value.
        GetCustomAttributesArgs[0] = ObjToArgSlot(MemberInfoObj);

        // Retrieve the custom attributes of type LCIDConversionAttribute.
        PTRARRAYREF CustomAttrArray = NULL;
        EX_TRY
        {
            CustomAttrArray = (PTRARRAYREF) getCustomAttributes.Call_RetOBJECTREF(GetCustomAttributesArgs);
        }
        EX_CATCH
        {
        }
        EX_END_CATCH

        GCPROTECT_BEGIN(CustomAttrArray)
        {
            if ((CustomAttrArray != NULL) && (CustomAttrArray->GetNumComponents() > 0))
                m_CultureAwareState = Aware;
            else
                m_CultureAwareState = NonAware;
        }
        GCPROTECT_END();
    }
    GCPROTECT_END();
}

void DispatchMemberInfo::SetUpParamMarshalerInfo()
{
    CONTRACTL
    {
        THROWS;
        GC_TRIGGERS;
        MODE_COOPERATIVE;
    }
    CONTRACTL_END;

    BOOL bSetUpReturnValueOnly = FALSE;
    OBJECTREF SetterObj = NULL;
    OBJECTREF GetterObj = NULL;
    OBJECTREF MemberInfoObj = GetMemberInfoObject();

    GCPROTECT_BEGIN(SetterObj);
    GCPROTECT_BEGIN(GetterObj);
    GCPROTECT_BEGIN(MemberInfoObj);
    {
        MethodTable *pMemberInfoMT = MemberInfoObj->GetMethodTable();

        if (CoreLibBinder::IsClass(pMemberInfoMT, CLASS__METHOD))
        {
            MethodDescCallSite getMethodHandle(METHOD__METHOD_BASE__GET_METHODDESC, &MemberInfoObj);
            ARG_SLOT arg = ObjToArgSlot(MemberInfoObj);
            MethodDesc* pMeth = (MethodDesc*) getMethodHandle.Call_RetLPVOID(&arg);
            if (pMeth)
                SetUpMethodMarshalerInfo(pMeth, FALSE);
        }
        else if (CoreLibBinder::IsClass(pMemberInfoMT, CLASS__FIELD))
        {
            // We don't support non-default marshalling behavior for field getter/setter stubs invoked via IDispatch.
        }
        else if (CoreLibBinder::IsClass(pMemberInfoMT, CLASS__PROPERTY))
        {
            BOOL isGetter = FALSE;
            MethodDescCallSite getSetter(METHOD__PROPERTY__GET_SETTER, &MemberInfoObj);
            ARG_SLOT args[] =
            {
                ObjToArgSlot(MemberInfoObj),
                BoolToArgSlot(false)
            };
            SetterObj = getSetter.Call_RetOBJECTREF(args);

            if (SetterObj != NULL)
            {
                MethodDescCallSite getMethodHandle(METHOD__METHOD_BASE__GET_METHODDESC, &SetterObj);
                ARG_SLOT arg = ObjToArgSlot(SetterObj);
                MethodDesc* pMeth = (MethodDesc*) getMethodHandle.Call_RetLPVOID(&arg);
                if (pMeth)
                {
                    bSetUpReturnValueOnly = TRUE;
                    SetUpMethodMarshalerInfo(pMeth, FALSE);
                }
            }

            MethodDescCallSite getGetter(METHOD__PROPERTY__GET_GETTER, &MemberInfoObj);
            ARG_SLOT args1[] =
            {
                ObjToArgSlot(MemberInfoObj),
                BoolToArgSlot(false)
            };
            GetterObj = getGetter.Call_RetOBJECTREF(args1);

            if (GetterObj != NULL)
            {
                MethodDescCallSite getMethodHandle(METHOD__METHOD_BASE__GET_METHODDESC, &GetterObj);
                ARG_SLOT arg = ObjToArgSlot(GetterObj);
                MethodDesc* pMeth = (MethodDesc*) getMethodHandle.Call_RetLPVOID(&arg);
                if (pMeth)
                {
                    // Only set up the marshalling information for the parameters if we
                    // haven't done it already for the setter.
                    SetUpMethodMarshalerInfo(pMeth, bSetUpReturnValueOnly);
                }
            }
        }
        else
        {
            // @FUTURE: Add support for user defined derived classes for
            //          MethodInfo, PropertyInfo and FieldInfo.
        }
    }
    GCPROTECT_END();
    GCPROTECT_END();
    GCPROTECT_END();
}

void DispatchMemberInfo::SetUpMethodMarshalerInfo(MethodDesc *pMD, BOOL bReturnValueOnly)
{
    CONTRACTL
    {
        THROWS;
        GC_TRIGGERS;
        MODE_ANY;
        PRECONDITION(CheckPointer(pMD));
        PRECONDITION(!pMD->IsAsyncMethod());
    }
    CONTRACTL_END;

    GCX_PREEMP();

    MetaSig         msig(pMD);
    LPCSTR          szName;
    USHORT          usSequence;
    DWORD           dwAttr;
    mdParamDef      returnParamDef = mdParamDefNil;
    mdParamDef      currParamDef = mdParamDefNil;

    int numArgs = msig.NumFixedArgs();

    IMDInternalImport *pInternalImport = msig.GetModule()->GetMDImport();

    HENUMInternalHolder hEnumParams(pInternalImport);

    //
    // Initialize the parameter definition enum.
    //
    hEnumParams.EnumInit(mdtParamDef, pMD->GetMemberDef());

    //
    // Retrieve the paramdef for the return type and determine which is the next
    // parameter that has parameter information.
    //
    do
    {
        if (pInternalImport->EnumNext(&hEnumParams, &currParamDef))
        {
            IfFailThrow(pInternalImport->GetParamDefProps(currParamDef, &usSequence, &dwAttr, &szName));

            if (usSequence == 0)
            {
                // The first parameter, if it has sequence 0, actually describes the return type.
                returnParamDef = currParamDef;
            }
        }
        else
        {
            usSequence = (USHORT)-1;
        }
    }
    while (usSequence == 0);

    // Look up the best fit mapping info via Assembly & Interface level attributes
    BOOL BestFit = TRUE;
    BOOL ThrowOnUnmappableChar = FALSE;
    ReadBestFitCustomAttribute(pMD, &BestFit, &ThrowOnUnmappableChar);

    //
    // Unless the bReturnValueOnly flag is set, set up the marshaling info for the parameters.
    //
    if (!bReturnValueOnly)
    {
        int iParam = 1;
        CorElementType  mtype;
        while (ELEMENT_TYPE_END != (mtype = msig.NextArg()))
        {
            //
            // Get the parameter token if the current parameter has one.
            //
            mdParamDef paramDef = mdParamDefNil;
            if (usSequence == iParam)
            {
                paramDef = currParamDef;

                if (pInternalImport->EnumNext(&hEnumParams, &currParamDef))
                {
                    IfFailThrow(pInternalImport->GetParamDefProps(currParamDef, &usSequence, &dwAttr, &szName));

                    // Validate that the param def tokens are in order.
                    _ASSERTE((usSequence > iParam) && "Param def tokens are not in order");
                }
                else
                {
                    usSequence = (USHORT)-1;
                }
            }


            //
            // Set up the marshaling info for the parameter.
            //

            MarshalInfo Info(msig.GetModule(), msig.GetArgProps(), msig.GetSigTypeContext(), paramDef, MarshalInfo::MARSHAL_SCENARIO_COMINTEROP,
                             (CorNativeLinkType)0, (CorNativeLinkFlags)0,
                             TRUE, iParam, numArgs, BestFit, ThrowOnUnmappableChar, FALSE, pMD, TRUE
    #ifdef _DEBUG
                     , pMD->m_pszDebugMethodName, pMD->m_pszDebugClassName, iParam
    #endif
                );


            //
            // Based on the MarshalInfo, set up a DispParamMarshaler for the parameter.
            //
            SetUpDispParamMarshalerForMarshalInfo(iParam, &Info);

            //
            // Get the in/out/ref attributes.
            //
            SetUpDispParamAttributes(iParam, &Info);

            m_bLastParamOleVarArg |= Info.IsOleVarArgCandidate();

            //
            // Increase the argument index.
            //
            iParam++;
        }

        // Make sure that there are not more param def tokens then there are CLR arguments.
        _ASSERTE( usSequence == (USHORT)-1 && "There are more parameter information tokens then there are CLR arguments" );
    }

    //
    // Set up the marshaling info for the return value.
    //

    if (!msig.IsReturnTypeVoid())
    {
        MarshalInfo Info(msig.GetModule(), msig.GetReturnProps(), msig.GetSigTypeContext(), returnParamDef, MarshalInfo::MARSHAL_SCENARIO_COMINTEROP,
                         (CorNativeLinkType)0, (CorNativeLinkFlags)0,
                         FALSE, 0, numArgs, BestFit, ThrowOnUnmappableChar, FALSE, pMD, TRUE
#ifdef _DEBUG
                         , pMD->m_pszDebugMethodName, pMD->m_pszDebugClassName, 0
#endif
                        );

        SetUpDispParamMarshalerForMarshalInfo(0, &Info);
    }
}

void DispatchMemberInfo::SetUpDispParamMarshalerForMarshalInfo(int iParam, MarshalInfo *pInfo)
{
    CONTRACTL
    {
        THROWS;
        GC_TRIGGERS;
        MODE_ANY;
        INJECT_FAULT(COMPlusThrowOM());
        PRECONDITION(CheckPointer(pInfo));
    }
    CONTRACTL_END;

    DispParamMarshaler *pDispParamMarshaler = pInfo->GenerateDispParamMarshaler();
    if (pDispParamMarshaler)
    {
        // If the array of marshalers hasn't been allocated yet, then allocate it.
        if (!m_apParamMarshaler)
        {
            // The array needs to be one more than the number of parameters for
            // normal methods and fields and 2 more properties.
            EnumMemberTypes MemberType = GetMemberType();
            int NumParamMarshalers = GetNumParameters() + ((MemberType == Property) ? 2 : 1);
            m_apParamMarshaler = new DispParamMarshaler*[NumParamMarshalers];
            memset(m_apParamMarshaler, 0, sizeof(DispParamMarshaler*) * NumParamMarshalers);
        }

        // Set the DispParamMarshaler in the array.
        m_apParamMarshaler[iParam] = pDispParamMarshaler;

        // If the disp param marshaler requires managed cleanup, then set
        // m_bRequiresManagedCleanup to TRUE to indicate the method requires
        // managed cleanup.
        if (pDispParamMarshaler->RequiresManagedCleanup())
            m_bRequiresManagedCleanup = TRUE;
    }
}


void DispatchMemberInfo::SetUpDispParamAttributes(int iParam, MarshalInfo* Info)
{
    CONTRACTL
    {
        THROWS;
        GC_TRIGGERS;
        MODE_ANY;
        INJECT_FAULT(COMPlusThrowOM());
        PRECONDITION(CheckPointer(Info));
    }
    CONTRACTL_END;

    // If the arry of In Only parameter indicators hasn't been allocated yet, then allocate it.
    if (!m_pParamInOnly)
    {
        // The array needs to be one more than the number of parameters for
        // normal methods and fields and 2 more properties.
        EnumMemberTypes MemberType = GetMemberType();
        int NumInOnlyFlags = GetNumParameters() + ((MemberType == Property) ? 2 : 1);
        m_pParamInOnly = new BOOL[NumInOnlyFlags];
        memset(m_pParamInOnly, 0, sizeof(BOOL) * NumInOnlyFlags);
    }

    m_pParamInOnly[iParam] = ( Info->IsIn() && !Info->IsOut() );
}

//--------------------------------------------------------------------------------
// The DispatchInfo class implementation.

DispatchInfo::DispatchInfo(MethodTable *pMT)
: m_pMT(pMT)
, m_pFirstMemberInfo(NULL)
, m_lock(CrstInterop, (CrstFlags)(CRST_REENTRANCY))
, m_CurrentDispID(0x10000)
, m_bAllowMembersNotInComMTMemberMap(FALSE)
, m_bInvokeUsingInvokeMember(FALSE)
{
    CONTRACTL
    {
        THROWS;
        GC_TRIGGERS;
        MODE_ANY;
        PRECONDITION(CheckPointer(pMT));
    }
    CONTRACTL_END;

    // Init the hashtable.
    m_DispIDToMemberInfoMap.Init(DISPID_TO_MEMBER_MAP_INITIAL_SIZE, NULL);
}

DispatchInfo::~DispatchInfo()
{
    CONTRACTL
    {
        NOTHROW;
        GC_NOTRIGGER;
        MODE_ANY;
    }
    CONTRACTL_END;

    DispatchMemberInfo* pCurrMember = m_pFirstMemberInfo;
    while (pCurrMember)
    {
        // Retrieve the next member.
        DispatchMemberInfo* pNextMember = pCurrMember->GetNext();

        // Delete the current member.
        delete pCurrMember;

        // Process the next member.
        pCurrMember = pNextMember;
    }
}

DispatchMemberInfo* DispatchInfo::FindMember(DISPID DispID)
{
    CONTRACT (DispatchMemberInfo*)
    {
        THROWS;
        GC_TRIGGERS;
        MODE_COOPERATIVE;
        POSTCONDITION(CheckPointer(RETVAL, NULL_OK));
    }
    CONTRACT_END;

    // We need to special case DISPID_UNKNOWN and -2 because the hashtable cannot handle them.
    // This is OK since these are invalid DISPID's.
    if ((DispID == DISPID_UNKNOWN) || (DispID == -2))
        RETURN NULL;

    // Lookup in the hashtable to find member with the specified DISPID. Note: this hash is unsynchronized, but Gethash
    // doesn't require synchronization.
    UPTR Data = (UPTR)m_DispIDToMemberInfoMap.Gethash(DispID2HashKey(DispID));
    if (Data != -1)
    {
        // We have found the member, so ensure it is initialized and return it.
        DispatchMemberInfo *pMemberInfo = (DispatchMemberInfo*)Data;

        pMemberInfo->EnsureInitialized();

        RETURN pMemberInfo;
    }
    else
    {
        RETURN NULL;
    }
}

DispatchMemberInfo* DispatchInfo::FindMember(SString& strName, BOOL bCaseSensitive)
{
    CONTRACT (DispatchMemberInfo*)
    {
        THROWS;
        GC_TRIGGERS;
        MODE_COOPERATIVE;
        POSTCONDITION(CheckPointer(RETVAL, NULL_OK));
    }
    CONTRACT_END;

    BOOL fFound = FALSE;

    // Go through the list of DispatchMemberInfo's to try and find one with the
    // specified name.
    DispatchMemberInfo *pCurrMemberInfo = m_pFirstMemberInfo;
    while (pCurrMemberInfo)
    {
        if (pCurrMemberInfo->GetMemberInfoObject() != NULL)
        {
            // Compare the 2 strings.
            SString& name = pCurrMemberInfo->GetName();
            if (bCaseSensitive
                    ? name.Equals(strName)
                    : name.EqualsCaseInsensitive(strName))
            {
                // We have found the member, so ensure it is initialized and return it.
                pCurrMemberInfo->EnsureInitialized();

                RETURN pCurrMemberInfo;
            }
        }

        // Process the next member.
        pCurrMemberInfo = pCurrMemberInfo->GetNext();
    }

    // No member has been found with the corresponding name.
    RETURN NULL;
}

// Helper method used to create DispatchMemberInfo's. This is only here because
// we can't call new inside a method that has a EX_TRY statement.
DispatchMemberInfo* DispatchInfo::CreateDispatchMemberInfoInstance(DISPID dispID, SString& strMemberName, OBJECTREF memberInfoObj)
{
    CONTRACT (DispatchMemberInfo*)
    {
        THROWS;
        GC_TRIGGERS;
        MODE_ANY;
        INJECT_FAULT(COMPlusThrowOM());
        POSTCONDITION(CheckPointer(RETVAL));
    }
    CONTRACT_END;

    DispatchMemberInfo* pInfo = new DispatchMemberInfo(this, dispID, strMemberName);
    pInfo->SetHandle(AllocateHandle(memberInfoObj));

    RETURN pInfo;
}

void DispatchInfo::InvokeMemberDebuggerWrapper(
                                      DispatchMemberInfo*   pDispMemberInfo,
                                      InvokeObjects*        pObjs,
                                      int                   NumParams,
                                      int                   NumArgs,
                                      int                   NumNamedArgs,
                                      int&                  iSrcArg,
                                      DISPID                id,
                                      DISPPARAMS*           pdp,
                                      VARIANT*              pVarRes,
                                      WORD                  wFlags,
                                      LCID                  lcid,
                                      DISPID*               pSrcArgNames,
                                      VARIANT*              pSrcArgs,
                                      Frame *               pFrame)

{
    // Use static contracts b/c we have SEH.
    STATIC_CONTRACT_THROWS;
    STATIC_CONTRACT_GC_TRIGGERS;
    STATIC_CONTRACT_MODE_ANY;

    // @todo - we have a PAL_TRY/PAL_EXCEPT here as a general (cross-platform) way to get a 1st-pass
    // filter. If that's bad perf, we could inline an FS:0 handler for x86-only; and then inline
    // both this wrapper and the main body.

    struct Param : public NotifyOfCHFFilterWrapperParam
    {
        DispatchInfo*         pThis;
        DispatchMemberInfo*   pDispMemberInfo;
        InvokeObjects*        pObjs;
        int                   NumParams;
        int                   NumArgs;
        int                   NumNamedArgs;
        int&                  iSrcArg;
        DISPID                id;
        DISPPARAMS*           pdp;
        VARIANT*              pVarRes;
        WORD                  wFlags;
        LCID                  lcid;
        DISPID*               pSrcArgNames;
        VARIANT*              pSrcArgs;

        Param(int& _iSrcArg)
            : iSrcArg(_iSrcArg)
        {}
    } param(iSrcArg);

    param.pFrame = GetThread()->GetFrame(); // Inherited from NotifyOfCHFFilterWrapperParam
    param.pThis = this;
    param.pDispMemberInfo = pDispMemberInfo;
    param.pObjs = pObjs;
    param.NumParams = NumParams;
    param.NumArgs = NumArgs;
    param.NumNamedArgs = NumNamedArgs;
    //param.iSrcArg = iSrcArg;
    param.id = id;
    param.pdp = pdp;
    param.pVarRes = pVarRes;
    param.wFlags = wFlags;
    param.lcid = lcid;
    param.pSrcArgNames = pSrcArgNames;
    param.pSrcArgs = pSrcArgs;

    PAL_TRY(Param *, pParam, &param)
    {
        struct
        {
            DispatchMemberInfo* pDispMemberInfo;
            BOOL* pParamInOnly;
            void(DispatchMemberInfo::*pCleanUpParamManaged)(int, OBJECTREF*);
            void(DispatchMemberInfo::*pMarshalParamNativeToManaged)(int, VARIANT*, OBJECTREF*);
            void(DispatchMemberInfo::*pMarshalParamManagedToNativeRef)(int, OBJECTREF*, VARIANT*);
            void(DispatchMemberInfo::*pMarshalReturnValueManagedToNative)(OBJECTREF*, VARIANT*);
            bool isLastParamOleVarArg;
            bool isCultureAware;
            bool requiresManagedObjCleanup;
        } dispMemberHelper;

        void* pDispMemberHelper = NULL;
        if (pParam->pDispMemberInfo)
        {
            dispMemberHelper.pDispMemberInfo = pParam->pDispMemberInfo;
            dispMemberHelper.pParamInOnly = pParam->pDispMemberInfo->m_pParamInOnly;
            dispMemberHelper.pCleanUpParamManaged = &DispatchMemberInfo::CleanUpParamManaged;
            dispMemberHelper.pMarshalParamNativeToManaged = &DispatchMemberInfo::MarshalParamNativeToManaged;
            dispMemberHelper.pMarshalParamManagedToNativeRef = &DispatchMemberInfo::MarshalParamManagedToNativeRef;
            dispMemberHelper.pMarshalReturnValueManagedToNative = &DispatchMemberInfo::MarshalReturnValueManagedToNative;
            dispMemberHelper.isLastParamOleVarArg = pParam->pDispMemberInfo->IsLastParamOleVarArg();
            dispMemberHelper.isCultureAware = pParam->pDispMemberInfo->IsCultureAware();
            dispMemberHelper.requiresManagedObjCleanup = pParam->pDispMemberInfo->RequiresManagedObjCleanup();
            pDispMemberHelper = &dispMemberHelper;
        }

        UnmanagedCallersOnlyCaller invokeMember(METHOD__DISPATCH_HELPER__INVOKE_MEMBER_WORKER);
        invokeMember.InvokeThrowing(
            pDispMemberHelper,
            &pParam->pObjs->ReflectionObj,
            &pParam->pObjs->MemberInfo,
            &pParam->pObjs->Target,
            pParam->NumParams,
            pParam->NumArgs,
            pParam->NumNamedArgs,
            &pParam->pSrcArgs,
            pParam->id,
            pParam->pdp,
            pParam->pVarRes,
            pParam->wFlags,
            pParam->lcid,
            pParam->pSrcArgNames,
            pParam->pSrcArgs);
    }
    PAL_EXCEPT_FILTER(NotifyOfCHFFilterWrapper)
    {
        // Should never reach here b/c handler should always continue search.
        _ASSERTE(false);
    }
    PAL_ENDTRY
}

// Helper method that invokes the member with the specified DISPID.
HRESULT DispatchInfo::InvokeMember(SimpleComCallWrapper *pSimpleWrap, DISPID id, LCID lcid, WORD wFlags, DISPPARAMS *pdp, VARIANT *pVarRes, EXCEPINFO *pei, IServiceProvider *pspCaller, unsigned int *puArgErr)
{
    CONTRACTL
    {
        THROWS;
        GC_TRIGGERS;
        MODE_COOPERATIVE;
        INJECT_FAULT(COMPlusThrowOM());
        PRECONDITION(CheckPointer(pSimpleWrap));
        PRECONDITION(CheckPointer(pdp, NULL_OK));
        PRECONDITION(CheckPointer(pVarRes, NULL_OK));
        PRECONDITION(CheckPointer(pei, NULL_OK));
        PRECONDITION(CheckPointer(pspCaller, NULL_OK));
        PRECONDITION(CheckPointer(puArgErr, NULL_OK));
    }
    CONTRACTL_END;

    HRESULT hr = S_OK;
    int iSrcArg = -1;
    int iBaseErrorArg = 0;
    int NumArgs;
    int NumNamedArgs;
    int NumParams;
    InvokeObjects Objs;
    DISPID *pSrcArgNames = NULL;
    VARIANT *pSrcArgs = NULL;
    ULONG_PTR ulActCtxCookie = 0;

    //
    // Validate the arguments.
    //

    if (!pdp)
        return E_POINTER;
    if (!pdp->rgvarg && (pdp->cArgs > 0))
        return E_INVALIDARG;
    if (!pdp->rgdispidNamedArgs && (pdp->cNamedArgs > 0))
        return E_INVALIDARG;
    if (pdp->cNamedArgs > pdp->cArgs)
        return E_INVALIDARG;
    if ((int)pdp->cArgs < 0 || (int)pdp->cNamedArgs < 0)
        return E_INVALIDARG;


    //
    // Clear the out arguments before we start.
    //

    if (pVarRes)
        SafeVariantClear(pVarRes);
    if (puArgErr)
        *puArgErr = -1;


    //
    // Convert the default LCID's to actual LCID's.
    //

    if(lcid == LOCALE_SYSTEM_DEFAULT || lcid == 0)
        lcid = GetSystemDefaultLCID();

    if(lcid == LOCALE_USER_DEFAULT)
        lcid = GetUserDefaultLCID();

    //
    // Set the value of the variables we use internally.
    //

    NumArgs = pdp->cArgs;
    NumNamedArgs = pdp->cNamedArgs;
    memset(&Objs, 0, sizeof(InvokeObjects));

    if (wFlags & (DISPATCH_PROPERTYPUT | DISPATCH_PROPERTYPUTREF))
    {
        // Since this invoke is for a property put or put ref we need to add 1 to
        // the iSrcArg to get the argument that is in error.
        iBaseErrorArg = 1;

        if (NumArgs < 1)
        {
            return DISP_E_BADPARAMCOUNT;
        }
        else
        {
            NumArgs--;
            pSrcArgs = &pdp->rgvarg[1];
        }

        if (NumNamedArgs < 1)
        {
            if (NumNamedArgs < 0)
                return DISP_E_BADPARAMCOUNT;

            // Verify if we really want to do this or return E_INVALIDARG instead.
            _ASSERTE(NumNamedArgs == 0);
            _ASSERTE(pSrcArgNames == NULL);
        }
        else
        {
            NumNamedArgs--;
            pSrcArgNames = &pdp->rgdispidNamedArgs[1];
        }
    }
    else
    {
        pSrcArgs = pdp->rgvarg;
        pSrcArgNames = pdp->rgdispidNamedArgs;
    }

    //
    // Do a lookup in the hashtable to find the DispatchMemberInfo for the DISPID.
    //

    DispatchMemberInfo *pDispMemberInfo = FindMember(id);
    if (!pDispMemberInfo || !pDispMemberInfo->GetMemberInfoObject())
    {
        pDispMemberInfo = NULL;
    }
    else if (pDispMemberInfo->IsNeutered())
    {
        COMPlusThrow(kInvalidOperationException);
    }

    //
    // If the member is not known then make sure that the DispatchInfo we have
    // supports unknown members.
    //

    if (m_bInvokeUsingInvokeMember)
    {
        // Since we do not have any information regarding the member then we
        // must assume the number of formal parameters matches the number of args.
        NumParams = NumArgs;
    }
    else
    {
        // If we haven't found the member then fail the invoke call.
        if (!pDispMemberInfo)
            return DISP_E_MEMBERNOTFOUND;

        // DISPATCH_CONSTRUCT only works when calling InvokeMember.
        if (wFlags & DISPATCH_CONSTRUCT)
            return E_INVALIDARG;

        if ((!(wFlags & (DISPATCH_METHOD | DISPATCH_PROPERTYGET))) && pDispMemberInfo->GetMemberType() == EnumMemberTypes::Method)
        {
            return DISP_E_MEMBERNOTFOUND;
        }

        // We have the member so retrieve the number of formal parameters.
        NumParams = pDispMemberInfo->GetNumParameters();

        if (pDispMemberInfo->IsLastParamOleVarArg())
        {
            // named args aren't allowed in a vararg function,
            // unless it's a lone DISPID_PROPERTYPUT (note that we already decrement
            // the value of NumNamedArgs for DISPID_PROPERTYPUT so that no special
            // check needed be done here for it
            // the logic is borrowed from the one in OLEAUT32!CTypeInfo2::Invoke
            if (NumNamedArgs > 0)
                return DISP_E_NONAMEDARGS;
        }
        else
        {
            // Make sure the number of arguments does not exceed the number of parameters.
            if(NumArgs > NumParams)
                return DISP_E_BADPARAMCOUNT;
        }

        // Validate that all the named arguments are known.
        for (iSrcArg = 0; iSrcArg < NumNamedArgs; iSrcArg++)
        {
            // There are some members we do not know about so we will call InvokeMember()
            // passing in the DISPID's directly so the caller can try to handle them.
            if (pSrcArgNames[iSrcArg] < 0 || pSrcArgNames[iSrcArg] >= NumParams)
                return DISP_E_MEMBERNOTFOUND;
        }
    }

    OBJECTREF pThrowable = NULL;

    //
    // The member is present so we need to convert the arguments and then do the
    // actual invocation.
    //
    GCPROTECT_BEGIN(pThrowable);
    GCPROTECT_BEGIN(Objs);
    {
        Objs.Target = pSimpleWrap->GetObjectRef();

        if (m_bInvokeUsingInvokeMember)
            Objs.ReflectionObj = GetReflectionObject();

        if (pDispMemberInfo)
            Objs.MemberInfo = pDispMemberInfo->GetMemberInfoObject();

        //
        // Invoke the method.
        //

        // The sole purpose of having this frame is to tell the debugger that we have a catch handler here
        // which may swallow managed exceptions.  The debugger needs this in order to send a
        // CatchHandlerFound (CHF) notification.
        DebuggerU2MCatchHandlerFrame catchFrame(true /* catchesAllExceptions */);

        EX_TRY
        {
            InvokeMemberDebuggerWrapper(pDispMemberInfo,
                                        &Objs,
                                        NumParams,
                                        NumArgs,
                                        NumNamedArgs,
                                        iSrcArg,
                                        id,
                                        pdp,
                                        pVarRes,
                                        wFlags,
                                        lcid,
                                        pSrcArgNames,
                                        pSrcArgs,
                                        &catchFrame);
        }
        EX_CATCH
        {
            pThrowable = GET_THROWABLE();
            RethrowTerminalExceptions();
        }
        EX_END_CATCH
        catchFrame.Pop();

        if (pThrowable != NULL)
        {
            // Do HR conversion.
            hr = SetupErrorInfo(pThrowable);
            if (hr == COR_E_TARGETINVOCATION)
            {
                hr = DISP_E_EXCEPTION;
                if (pei)
                {
                    // Retrieve the exception iformation.
                    GetExcepInfoForInvocationExcep(pThrowable, pei);

                    // Clear the IErrorInfo on the current thread since it does contains
                    // information on the TargetInvocationException which conflicts with
                    // the information in the returned EXCEPINFO.
                    IErrorInfo *pErrInfo = NULL;
                    HRESULT hr2 = SafeGetErrorInfo(&pErrInfo);
                    _ASSERTE(hr2 == S_OK);
                    SafeRelease(pErrInfo);
                }
            }
            else if (hr == COR_E_OVERFLOW)
            {
                hr = DISP_E_OVERFLOW;
                if (iSrcArg != -1)
                {
                    if (puArgErr)
                        *puArgErr = iSrcArg + iBaseErrorArg;
                }
            }
            else if (hr == COR_E_INVALIDOLEVARIANTTYPE)
            {
                hr = DISP_E_BADVARTYPE;
                if (iSrcArg != -1)
                {
                    if (puArgErr)
                        *puArgErr = iSrcArg + iBaseErrorArg;
                }
            }
            else if (hr == COR_E_ARGUMENT)
            {
                hr = E_INVALIDARG;
                if (iSrcArg != -1)
                {
                    if (puArgErr)
                        *puArgErr = iSrcArg + iBaseErrorArg;
                }
            }
            else if (hr == COR_E_SAFEARRAYTYPEMISMATCH)
            {
                hr = DISP_E_TYPEMISMATCH;
                if (iSrcArg != -1)
                {
                    if (puArgErr)
                        *puArgErr = iSrcArg + iBaseErrorArg;
                }
            }
            else if (hr == COR_E_MISSINGMEMBER || hr == COR_E_MISSINGMETHOD)
            {
                hr = DISP_E_MEMBERNOTFOUND;

                // This exception should never occur while we are marshaling arguments.
                _ASSERTE(iSrcArg == -1);
            }
        }
    }
    GCPROTECT_END();
    GCPROTECT_END();
    return hr;
}

// Parameter marshaling helpers.
void DispatchInfo::MarshalParamNativeToManaged(DispatchMemberInfo *pMemberInfo, int iParam, VARIANT *pSrcVar, OBJECTREF *pDestObj)
{
    CONTRACTL
    {
        THROWS;
        GC_TRIGGERS;
        MODE_COOPERATIVE;
    }
    CONTRACTL_END;

    if (pMemberInfo && !m_bInvokeUsingInvokeMember)
        pMemberInfo->MarshalParamNativeToManaged(iParam, pSrcVar, pDestObj);
    else
        OleVariant::MarshalObjectForOleVariant(pSrcVar, pDestObj);
}

void DispatchInfo::MarshalParamManagedToNativeRef(DispatchMemberInfo *pMemberInfo, int iParam, OBJECTREF *pSrcObj, OBJECTREF *pBackupStaticArray, VARIANT *pRefVar)
{
    CONTRACTL
    {
        THROWS;
        GC_TRIGGERS;
        MODE_COOPERATIVE;
        PRECONDITION(CheckPointer(pMemberInfo, NULL_OK));
        PRECONDITION(pSrcObj != NULL);
        PRECONDITION(CheckPointer(pRefVar));
    }
    CONTRACTL_END;

    if (pBackupStaticArray && (*pBackupStaticArray != NULL))
    {
        // The contents of a static array can change, but not the array itself. If
        // the array has changed, then throw an exception.
        if (*pSrcObj != *pBackupStaticArray)
            COMPlusThrow(kInvalidOperationException, IDS_INVALID_REDIM);

        // Retrieve the element VARTYPE and method table.
        VARTYPE ElementVt = V_VT(pRefVar) & ~(VT_BYREF | VT_ARRAY);
        MethodTable *pElementMT = (*(BASEARRAYREF *)pSrcObj)->GetArrayElementTypeHandle().GetMethodTable();

        PCODE pStructMarshalStubAddress = NULL;
        GCPROTECT_BEGIN(*pSrcObj);
        if (ElementVt == VT_RECORD && pElementMT->IsBlittable())
        {
            GCX_PREEMP();
            pStructMarshalStubAddress = PInvoke::GetEntryPointForStructMarshalStub(pElementMT);
        }
        GCPROTECT_END();

        // Convert the contents of the managed array into the original SAFEARRAY.
        OleVariant::MarshalSafeArrayForArrayRef((BASEARRAYREF *)pSrcObj, *V_ARRAYREF(pRefVar), ElementVt, pElementMT, pStructMarshalStubAddress);
    }
    else
{
    if (pMemberInfo && !m_bInvokeUsingInvokeMember)
        pMemberInfo->MarshalParamManagedToNativeRef(iParam, pSrcObj, pRefVar);
    else
        OleVariant::MarshalOleRefVariantForObject(pSrcObj, pRefVar);
}
}

void DispatchInfo::MarshalReturnValueManagedToNative(DispatchMemberInfo *pMemberInfo, OBJECTREF *pSrcObj, VARIANT *pDestVar)
{
    CONTRACTL
    {
        THROWS;
        GC_TRIGGERS;
        MODE_COOPERATIVE;
    }
    CONTRACTL_END;

    if (pMemberInfo && !m_bInvokeUsingInvokeMember)
        pMemberInfo->MarshalReturnValueManagedToNative(pSrcObj, pDestVar);
    else
        OleVariant::MarshalOleVariantForObject(pSrcObj, pDestVar);
}

void DispatchInfo::CleanUpNativeParam(DispatchMemberInfo *pDispMemberInfo, int iParamIndex, OBJECTREF *pBackupStaticArray, VARIANT *pArgVariant)
{
    CONTRACTL
    {
        NOTHROW;
        GC_TRIGGERS;
        MODE_COOPERATIVE;
        PRECONDITION(pArgVariant != NULL);
    }
    CONTRACTL_END;

    EX_TRY
    {
        switch (V_VT(pArgVariant) & ~VT_BYREF)
        {
            case VT_I1:    case VT_I2:    case VT_I4:    case VT_I8:
            case VT_UI1:   case VT_UI2:   case VT_UI4:   case VT_UI8:
            case VT_INT:   case VT_UINT:  case VT_PTR:
            case VT_R4:    case VT_R8:    case VT_BOOL:
            case VT_CY:    case VT_DATE:
            case VT_ERROR: case VT_HRESULT:
            case VT_DECIMAL:
            {
                // the argument type is a value type - overwrite it with zeros
                UINT uSize = OleVariant::GetElementSizeForVarType(V_VT(pArgVariant) & ~VT_BYREF, NULL);
                FillMemory(V_BYREF(pArgVariant), uSize, 0);
                break;
            }

            default:
            {
                // marshal managed null into the VARIANT which works for reference types
                OBJECTREF Null = NULL;

                GCPROTECT_BEGIN(Null); // the local stays NULL, this is just to satisfy contracts
                MarshalParamManagedToNativeRef(pDispMemberInfo, iParamIndex, &Null, pBackupStaticArray, pArgVariant);
                GCPROTECT_END();
            }
        }
    }
    EX_CATCH
    {
        // if the argument was totally corrupted and cleanup failed, just swallow it and continue
    }
    EX_END_CATCH
}

MethodDesc* DispatchInfo::GetFieldInfoMD(BinderMethodID Method, TypeHandle hndFieldInfoType)
{
    CONTRACT (MethodDesc*)
    {
        THROWS;
        GC_TRIGGERS;
        MODE_ANY;
        POSTCONDITION(CheckPointer(RETVAL));
    }
    CONTRACT_END

    MethodDesc *pMD;

    // If the current class is the standard implementation then return the cached method desc
    if (CoreLibBinder::IsClass(hndFieldInfoType.GetMethodTable(), CLASS__FIELD))
    {
        pMD = CoreLibBinder::GetMethod(Method);
    }
    else
    {
        pMD = MemberLoader::FindMethod(hndFieldInfoType.GetMethodTable(),
                CoreLibBinder::GetMethodName(Method), CoreLibBinder::GetMethodSig(Method));
    }
    _ASSERTE(pMD && "Unable to find specified FieldInfo method");

    // Return the specified method desc.
    RETURN pMD;
}

MethodDesc* DispatchInfo::GetPropertyInfoMD(BinderMethodID Method, TypeHandle hndPropInfoType)
{
    CONTRACT (MethodDesc*)
    {
        THROWS;
        GC_TRIGGERS;
        MODE_ANY;
        POSTCONDITION(CheckPointer(RETVAL));
    }
    CONTRACT_END

    MethodDesc *pMD;

    // If the current class is the standard implementation then return the cached method desc if present.
    if (CoreLibBinder::IsClass(hndPropInfoType.GetMethodTable(), CLASS__PROPERTY))
    {
        pMD = CoreLibBinder::GetMethod(Method);
    }
    else
    {
        pMD = MemberLoader::FindMethod(hndPropInfoType.GetMethodTable(),
                CoreLibBinder::GetMethodName(Method), CoreLibBinder::GetMethodSig(Method));
    }
    _ASSERTE(pMD && "Unable to find specified PropertyInfo method");

    // Return the specified method desc.
    RETURN pMD;
}

MethodDesc* DispatchInfo::GetMethodInfoMD(BinderMethodID Method, TypeHandle hndMethodInfoType)
{
    CONTRACT (MethodDesc*)
    {
        THROWS;
        GC_TRIGGERS;
        MODE_ANY;
        POSTCONDITION(CheckPointer(RETVAL));
    }
    CONTRACT_END

    MethodDesc *pMD;

    // If the current class is the standard implementation then return the cached method desc.
    if (CoreLibBinder::IsClass(hndMethodInfoType.GetMethodTable(), CLASS__METHOD))
    {
        pMD = CoreLibBinder::GetMethod(Method);
    }
    else
    {
        pMD = MemberLoader::FindMethod(hndMethodInfoType.GetMethodTable(),
                CoreLibBinder::GetMethodName(Method), CoreLibBinder::GetMethodSig(Method));
    }
    _ASSERTE(pMD && "Unable to find specified MethodInfo method");

    // Return the specified method desc.
    RETURN pMD;
}

MethodDesc* DispatchInfo::GetCustomAttrProviderMD(TypeHandle hndCustomAttrProvider)
{
    CONTRACT (MethodDesc*)
    {
        THROWS;
        GC_TRIGGERS;
        MODE_ANY;
        POSTCONDITION(CheckPointer(RETVAL));
    }
    CONTRACT_END;

    MethodTable *pMT = hndCustomAttrProvider.AsMethodTable();
    MethodDesc *pMD = pMT->GetMethodDescForInterfaceMethod(CoreLibBinder::GetMethod(METHOD__ICUSTOM_ATTR_PROVIDER__GET_CUSTOM_ATTRIBUTES), TRUE /* throwOnConflict */);

    // Return the specified method desc.
    RETURN pMD;
}

// This method synchronizes the DispatchInfo's members with the ones in the method tables type.
// The return value will be set to TRUE if the object was out of synch and members where
// added and it will be set to FALSE otherwise.
BOOL DispatchInfo::SynchWithManagedView()
{
    CONTRACTL
    {
        THROWS;
        GC_TRIGGERS;
        MODE_ANY;
    }
    CONTRACTL_END;

    HRESULT hr = S_OK;
    NewArrayHolder<WCHAR> strMemberName = NULL;
    NewHolder<ComMTMemberInfoMap> pMemberMap = NULL;

    // This represents the new member to add and it is also used to determine if members have
    // been added or not.
    NewHolder<DispatchMemberInfo> pMemberToAdd = NULL;

    Thread* pThread = SetupThreadNoThrow();
    if (pThread == NULL)
        return FALSE;

    // Determine if this is the first time we synch.
    BOOL bFirstSynch = (m_pFirstMemberInfo == NULL);

    // This method needs to be synchronized to make sure two threads don't try and
    // add members at the same time.
    CrstHolder ch(&m_lock);
    {
        // Make sure we switch to cooperative mode before we start.
        GCX_COOP();

        // Go through the list of member info's and find the end.
        DispatchMemberInfo **ppNextMember = &m_pFirstMemberInfo;
        while (*ppNextMember)
            ppNextMember = (*ppNextMember)->GetNextPtr();

        // Retrieve the member info map.
        pMemberMap = GetMemberInfoMap();

        for (int cPhase = 0; cPhase < 3; cPhase++)
        {
            PTRARRAYREF MemberArrayObj = NULL;
            GCPROTECT_BEGIN(MemberArrayObj);

            // Retrieve the appropriate array of members for the current phase.
            switch (cPhase)
            {
                case 0:
                    // Retrieve the array of properties.
                    MemberArrayObj = RetrievePropList();
                    break;

                case 1:
                    // Retrieve the array of fields.
                    MemberArrayObj = RetrieveFieldList();
                    break;

                case 2:
                    // Retrieve the array of methods.
                    MemberArrayObj = RetrieveMethList();
                    break;
            }

            // Retrieve the number of components in the member array.
            UINT NumComponents = 0;
            if (MemberArrayObj != NULL)
                NumComponents = MemberArrayObj->GetNumComponents();

            // Go through all the member info's in the array and see if they are already
            // in the DispatchExInfo.
            for (UINT i = 0; i < NumComponents; i++)
            {
                BOOL bMatch = FALSE;

                OBJECTREF CurrMemberInfoObj = MemberArrayObj->GetAt(i);
                GCPROTECT_BEGIN(CurrMemberInfoObj)
                {
                    DispatchMemberInfo *pCurrMemberInfo = m_pFirstMemberInfo;
                    while (pCurrMemberInfo)
                    {
                        // We can simply compare the OBJECTREF's.
                        if (CurrMemberInfoObj == pCurrMemberInfo->GetMemberInfoObject())
                        {
                            // We have found a match.
                            bMatch = TRUE;
                            break;
                        }

                        // Check the next member.
                        pCurrMemberInfo = pCurrMemberInfo->GetNext();
                    }

                    // If we have not found a match then we need to add the member info to the
                    // list of member info's that will be added to the DispatchExInfo.
                    if (!bMatch)
                    {
                        DISPID MemberID = DISPID_UNKNOWN;
                        BOOL bAddMember = FALSE;


                        //
                        // Attempt to retrieve the properties of the member.
                        //

                        ComMTMethodProps *pMemberProps = DispatchMemberInfo::GetMemberProps(CurrMemberInfoObj, pMemberMap);

                        //
                        // Determine if we are to add this member or not.
                        //

                        if (pMemberProps)
                            bAddMember = pMemberProps->bMemberVisible;
                        else
                            bAddMember = m_bAllowMembersNotInComMTMemberMap;

                        if (bAddMember)
                        {
                            //
                            // Retrieve the DISPID of the member.
                            //
                            MemberID = DispatchMemberInfo::GetMemberDispId(CurrMemberInfoObj, pMemberMap);

                            //
                            // If the member does not have an explicit DISPID or if the specified DISPID
                            // is already in use then we need to generate a dynamic DISPID for the member.
                            //

                            if ((MemberID == DISPID_UNKNOWN) || (FindMember(MemberID) != NULL))
                                MemberID = GenerateDispID();

                            //
                            // Retrieve the name of the member.
                            //

                            strMemberName = DispatchMemberInfo::GetMemberName(CurrMemberInfoObj, pMemberMap);

                            //
                            // Create a DispatchInfoMemberInfo that will represent the member.
                            //

                            SString sName(strMemberName);
                            pMemberToAdd = CreateDispatchMemberInfoInstance(MemberID, sName, CurrMemberInfoObj);

                            //
                            // Add the member to the end of the list.
                            //

                            *ppNextMember = pMemberToAdd;

                            // Update ppNextMember to be ready for the next new member.
                            ppNextMember = (*ppNextMember)->GetNextPtr();

                            // Add the member to the map. Note, the hash is unsynchronized, but we already have our lock
                            // so we're okay.
                            m_DispIDToMemberInfoMap.InsertValue(DispID2HashKey(MemberID), pMemberToAdd);
                            pMemberToAdd.SuppressRelease();
                        }
                    }
                }
                GCPROTECT_END();
            }

            GCPROTECT_END();
        }
        // GC mode toggles back here
    }
    // Check to see if any new members were added to the expando object.
    return pMemberToAdd ? TRUE : FALSE;

    // locks released and memory cleaned up here
}

BOOL DispatchInfo::VariantIsMissing(VARIANT *pOle)
{
    LIMITED_METHOD_CONTRACT;

    return (V_VT(pOle) == VT_ERROR) && (V_ERROR(pOle) == DISP_E_PARAMNOTFOUND);
}

LOADERHANDLE DispatchInfo::AllocateHandle(OBJECTREF objRef)
{
    WRAPPER_NO_CONTRACT;

    return m_pMT->GetLoaderAllocator()->AllocateHandle(objRef);
}

void DispatchInfo::FreeHandle(LOADERHANDLE handle)
{
    CONTRACTL
    {
        NOTHROW;
        GC_NOTRIGGER;
        MODE_ANY;
        PRECONDITION(handle != NULL);
    }
    CONTRACTL_END;

    PTR_LoaderAllocator loaderAllocator = m_pMT->GetLoaderAllocator();

    // If the loader isn't alive, we can't free the handle.
    if (loaderAllocator->AddReferenceIfAlive())
    {
        loaderAllocator->FreeHandle(handle);
        loaderAllocator->Release();
    }
}

OBJECTREF DispatchInfo::GetHandleValue(LOADERHANDLE handle)
{
    WRAPPER_NO_CONTRACT;

    return m_pMT->GetLoaderAllocator()->GetHandleValue(handle);
}

PTRARRAYREF DispatchInfo::RetrievePropList()
{
    CONTRACTL
    {
        THROWS;
        GC_TRIGGERS;
        MODE_COOPERATIVE;
    }
    CONTRACTL_END;

    // return value
    PTRARRAYREF orRetVal;

    // Retrieve the exposed class object.
    OBJECTREF TargetObj = GetReflectionObject();

    GCPROTECT_BEGIN(TargetObj);
    MethodDescCallSite getProperties(METHOD__CLASS__GET_PROPERTIES, &TargetObj);

    // Prepare the arguments that will be passed to the method.
    ARG_SLOT Args[] =
    {
        ObjToArgSlot(TargetObj),
        (ARG_SLOT)BINDER_DefaultLookup
    };

    // Retrieve the array of members from the type object.
    orRetVal = (PTRARRAYREF) getProperties.Call_RetOBJECTREF(Args);

    GCPROTECT_END();

    return orRetVal;
}

PTRARRAYREF DispatchInfo::RetrieveFieldList()
{
    CONTRACTL
    {
        THROWS;
        GC_TRIGGERS;
        MODE_COOPERATIVE;
    }
    CONTRACTL_END;

    // return value
    PTRARRAYREF orRetVal;

    // Retrieve the exposed class object.
    OBJECTREF TargetObj = GetReflectionObject();

    GCPROTECT_BEGIN(TargetObj);
    MethodDescCallSite getFields(METHOD__CLASS__GET_FIELDS, &TargetObj);

    // Prepare the arguments that will be passed to the method.
    ARG_SLOT Args[] =
    {
        ObjToArgSlot(TargetObj),
        (ARG_SLOT)BINDER_DefaultLookup
    };

    // Retrieve the array of members from the type object.
    orRetVal = (PTRARRAYREF) getFields.Call_RetOBJECTREF(Args);

    GCPROTECT_END();

    return orRetVal;
}

PTRARRAYREF DispatchInfo::RetrieveMethList()
{
    CONTRACTL
    {
        THROWS;
        GC_TRIGGERS;
        MODE_COOPERATIVE;
    }
    CONTRACTL_END;

    // return value
    PTRARRAYREF orRetVal;

    // Retrieve the exposed class object.
    OBJECTREF TargetObj = GetReflectionObject();

    GCPROTECT_BEGIN(TargetObj);
    MethodDescCallSite getMethods(METHOD__CLASS__GET_METHODS, &TargetObj);

    // Prepare the arguments that will be passed to the method.
    ARG_SLOT Args[] =
    {
        ObjToArgSlot(TargetObj),
        (ARG_SLOT)BINDER_DefaultLookup
    };

    // Retrieve the array of members from the type object.
    orRetVal = (PTRARRAYREF) getMethods.Call_RetOBJECTREF(Args);

    GCPROTECT_END();

    return orRetVal;
}

// Virtual method to retrieve the InvokeMember method desc.
MethodDesc* DispatchInfo::GetInvokeMemberMD()
{
    CONTRACT (MethodDesc*)
    {
        THROWS;
        GC_TRIGGERS;
        MODE_ANY;
        POSTCONDITION(CheckPointer(RETVAL));
    }
    CONTRACT_END;

    RETURN CoreLibBinder::GetMethod(METHOD__CLASS__INVOKE_MEMBER);
}

// Virtual method to retrieve the object associated with this DispatchInfo that
// implements IReflect.
OBJECTREF DispatchInfo::GetReflectionObject()
{
    CONTRACTL
    {
        THROWS;
        GC_TRIGGERS;
        MODE_COOPERATIVE;
    }
    CONTRACTL_END;

    return m_pMT->GetManagedClassObject();
}

// Virtual method to retrieve the member info map.
ComMTMemberInfoMap *DispatchInfo::GetMemberInfoMap()
{
    CONTRACT (ComMTMemberInfoMap*)
    {
        THROWS;
        GC_TRIGGERS;
        MODE_COOPERATIVE;
        INJECT_FAULT(COMPlusThrowOM());
        POSTCONDITION(CheckPointer(RETVAL));
    }
    CONTRACT_END;


    // Create the member info map.
    NewHolder<ComMTMemberInfoMap> pMemberInfoMap (new ComMTMemberInfoMap(m_pMT));

    // Initialize it.
    pMemberInfoMap->Init(sizeof(void*));

    pMemberInfoMap.SuppressRelease();
    RETURN pMemberInfoMap;
}

// Helper function to fill in an EXCEPINFO for an InvocationException.
void DispatchInfo::GetExcepInfoForInvocationExcep(OBJECTREF objException, EXCEPINFO *pei)
{
    CONTRACTL
    {
        THROWS;
        GC_TRIGGERS;
        MODE_COOPERATIVE;
        PRECONDITION(objException != NULL);
        PRECONDITION(CheckPointer(pei));
    }
    CONTRACTL_END;

    MethodDesc *pMD;
    ExceptionData ED;
    OBJECTREF InnerExcep = NULL;

    // Initialize the EXCEPINFO.
    memset(pei, 0, sizeof(EXCEPINFO));
    pei->scode = E_FAIL;

    GCPROTECT_BEGIN(InnerExcep)
    GCPROTECT_BEGIN(objException)
    {
        // Retrieve the method desc to access the InnerException property.
        pMD = MemberLoader::FindPropertyMethod(objException->GetMethodTable(), EXCEPTION_INNER_PROP, PropertyGet);
        _ASSERTE(pMD && "Unable to find get method for proprety Exception.InnerException");
        MethodDescCallSite propGet(pMD, &objException);

        // Retrieve the value of the InnerException property.
        ARG_SLOT GetInnerExceptionArgs[] = { ObjToArgSlot(objException) };
        InnerExcep = propGet.Call_RetOBJECTREF(GetInnerExceptionArgs);

        // If the inner exception object is null then we can't get any info.
        if (InnerExcep != NULL)
        {
            // Retrieve the exception data for the inner exception.
            ExceptionNative::GetExceptionData(InnerExcep, &ED);
            pei->bstrSource = ED.bstrSource;
            pei->bstrDescription = ED.bstrDescription;
            pei->bstrHelpFile = ED.bstrHelpFile;
            pei->dwHelpContext = ED.dwHelpContext;
            pei->scode = ED.hr;
        }
    }
    GCPROTECT_END();
    GCPROTECT_END();
}

DISPID DispatchInfo::GenerateDispID()
{
    CONTRACTL
    {
        NOTHROW;
        GC_NOTRIGGER;
        MODE_ANY;
    }
    CONTRACTL_END;

    // Find the next unused DISPID. Note, the hash is unsynchronized, but Gethash doesn't require synchronization.
    for (; (UPTR)m_DispIDToMemberInfoMap.Gethash(DispID2HashKey(m_CurrentDispID)) != -1; m_CurrentDispID++);
    return m_CurrentDispID++;
}

//--------------------------------------------------------------------------------
// The DispatchExInfo class implementation.

DispatchExInfo::DispatchExInfo(SimpleComCallWrapper *pSimpleWrapper, MethodTable *pMT)
: DispatchInfo(pMT)
, m_pSimpleWrapperOwner(pSimpleWrapper)
{
    CONTRACTL
    {
        THROWS;
        GC_TRIGGERS;
        MODE_ANY;
        PRECONDITION(CheckPointer(pSimpleWrapper));
        PRECONDITION(CheckPointer(pMT));
    }
    CONTRACTL_END;

    // Set the flags to specify the behavior of the base DispatchInfo class.
    m_bAllowMembersNotInComMTMemberMap = TRUE;
    m_bInvokeUsingInvokeMember = TRUE;
}

DispatchExInfo::~DispatchExInfo()
{
    WRAPPER_NO_CONTRACT;
}

// Methods to lookup members. These methods synch with the managed view if they fail to
// find the method.
DispatchMemberInfo* DispatchExInfo::SynchFindMember(DISPID DispID)
{
    CONTRACT (DispatchMemberInfo*)
    {
        THROWS;
        GC_TRIGGERS;
        MODE_ANY;
        POSTCONDITION(CheckPointer(RETVAL, NULL_OK));
    }
    CONTRACT_END;

    DispatchMemberInfo *pMemberInfo = FindMember(DispID);

    if (!pMemberInfo && SynchWithManagedView())
        pMemberInfo = FindMember(DispID);

    RETURN pMemberInfo;
}

DispatchMemberInfo* DispatchExInfo::SynchFindMember(SString& strName, BOOL bCaseSensitive)
{
    CONTRACT (DispatchMemberInfo*)
    {
        THROWS;
        GC_TRIGGERS;
        MODE_ANY;
        POSTCONDITION(CheckPointer(RETVAL, NULL_OK));
    }
    CONTRACT_END;

    DispatchMemberInfo *pMemberInfo = FindMember(strName, bCaseSensitive);

    if (!pMemberInfo && SynchWithManagedView())
        pMemberInfo = FindMember(strName, bCaseSensitive);

    RETURN pMemberInfo;
}

// Helper method that invokes the member with the specified DISPID. These methods synch
// with the managed view if they fail to find the method.
HRESULT DispatchExInfo::SynchInvokeMember(SimpleComCallWrapper *pSimpleWrap, DISPID id, LCID lcid, WORD wFlags, DISPPARAMS *pdp, VARIANT *pVarRes, EXCEPINFO *pei, IServiceProvider *pspCaller, unsigned int *puArgErr)
{
    CONTRACTL
    {
        THROWS;
        GC_TRIGGERS;
        MODE_COOPERATIVE;
    }
    CONTRACTL_END;

    // Invoke the member.
    HRESULT hr = InvokeMember(pSimpleWrap, id, lcid, wFlags, pdp, pVarRes, pei, pspCaller, puArgErr);

    // If the member was not found then we need to synch and try again if the managed view has changed.
    if ((hr == DISP_E_MEMBERNOTFOUND) && SynchWithManagedView())
        hr = InvokeMember(pSimpleWrap, id, lcid, wFlags, pdp, pVarRes, pei, pspCaller, puArgErr);

    return hr;
}

DispatchMemberInfo* DispatchExInfo::GetFirstMember()
{
    CONTRACT (DispatchMemberInfo*)
    {
        THROWS;
        GC_TRIGGERS;
        MODE_COOPERATIVE;
        POSTCONDITION(CheckPointer(RETVAL, NULL_OK));
    }
    CONTRACT_END;

    // Start with the first member.
    DispatchMemberInfo **ppNextMemberInfo = &m_pFirstMemberInfo;

    // If the next member is not set we need to sink up with the expando object
    // itself to make sure that this member is really the last member and that
    // other members have not been added without us knowing.
    if (!(*ppNextMemberInfo))
    {
        if (SynchWithManagedView())
        {
            // New members have been added to the list and since they must be added
            // to the end the next member of the previous end of the list must
            // have been updated.
            _ASSERTE(*ppNextMemberInfo);
        }
    }

    // Now we need to make sure we skip any members that are deleted.
    while ((*ppNextMemberInfo) && !(*ppNextMemberInfo)->GetMemberInfoObject())
        ppNextMemberInfo = (*ppNextMemberInfo)->GetNextPtr();

    RETURN *ppNextMemberInfo;
}

DispatchMemberInfo* DispatchExInfo::GetNextMember(DISPID CurrMemberDispID)
{
    CONTRACT (DispatchMemberInfo*)
    {
        THROWS;
        GC_TRIGGERS;
        MODE_COOPERATIVE;
        POSTCONDITION(CheckPointer(RETVAL, NULL_OK));
    }
    CONTRACT_END;

    // Do a lookup in the hashtable to find the DispatchMemberInfo for the DISPID.
    DispatchMemberInfo *pDispMemberInfo = FindMember(CurrMemberDispID);
    if (!pDispMemberInfo)
        RETURN NULL;

    // Start from the next member.
    DispatchMemberInfo **ppNextMemberInfo = pDispMemberInfo->GetNextPtr();

    // If the next member is not set we need to sink up with the expando object
    // itself to make sure that this member is really the last member and that
    // other members have not been added without us knowing.
    if (!(*ppNextMemberInfo))
    {
        if (SynchWithManagedView())
        {
            // New members have been added to the list and since they must be added
            // to the end the next member of the previous end of the list must
            // have been updated.
            _ASSERTE(*ppNextMemberInfo);
        }
    }

    // Now we need to make sure we skip any members that are deleted.
    while ((*ppNextMemberInfo) && !(*ppNextMemberInfo)->GetMemberInfoObject())
        ppNextMemberInfo = (*ppNextMemberInfo)->GetNextPtr();

    RETURN *ppNextMemberInfo;
}

MethodDesc* DispatchExInfo::GetIReflectMD(BinderMethodID Method)
{
    CONTRACT (MethodDesc*)
    {
        THROWS;
        GC_TRIGGERS;
        MODE_ANY;
        POSTCONDITION(CheckPointer(RETVAL));
    }
    CONTRACT_END;

    MethodTable *pMT = m_pSimpleWrapperOwner->GetMethodTable();
    MethodDesc *pMD = pMT->GetMethodDescForInterfaceMethod(CoreLibBinder::GetMethod(Method), TRUE /* throwOnConflict */);

    // Return the specified method desc.
    RETURN pMD;
}

PTRARRAYREF DispatchExInfo::RetrievePropList()
{
    CONTRACTL
    {
        THROWS;
        GC_TRIGGERS;
        MODE_COOPERATIVE;
    }
    CONTRACTL_END;

    PTRARRAYREF oPropList;

    // Retrieve the expando OBJECTREF.
    OBJECTREF TargetObj = GetReflectionObject();
    GCPROTECT_BEGIN(TargetObj);

    // Retrieve the GetMembers MethodDesc.
    MethodDesc *pMD = GetIReflectMD(METHOD__IREFLECT__GET_PROPERTIES);
    MethodDescCallSite getProperties(pMD, &TargetObj);

    // Prepare the arguments that will be passed to the method.
    ARG_SLOT Args[] =
    {
        ObjToArgSlot(TargetObj),
        (ARG_SLOT)BINDER_DefaultLookup
    };

    // Retrieve the array of members from the expando object
    oPropList = (PTRARRAYREF) getProperties.Call_RetOBJECTREF(Args);

    GCPROTECT_END();

    return oPropList;
}

PTRARRAYREF DispatchExInfo::RetrieveFieldList()
{
    CONTRACTL
    {
        THROWS;
        GC_TRIGGERS;
        MODE_COOPERATIVE;
    }
    CONTRACTL_END;

    PTRARRAYREF oFieldList;

    // Retrieve the expando OBJECTREF.
    OBJECTREF TargetObj = GetReflectionObject();
    GCPROTECT_BEGIN(TargetObj);

    // Retrieve the GetMembers MethodDesc.
    MethodDesc *pMD = GetIReflectMD(METHOD__IREFLECT__GET_FIELDS);
    MethodDescCallSite getFields(pMD, &TargetObj);

    // Prepare the arguments that will be passed to the method.
    ARG_SLOT Args[] =
    {
        ObjToArgSlot(TargetObj),
        (ARG_SLOT)BINDER_DefaultLookup
    };

    // Retrieve the array of members from the expando object
    oFieldList = (PTRARRAYREF) getFields.Call_RetOBJECTREF(Args);

    GCPROTECT_END();

    return oFieldList;
}

PTRARRAYREF DispatchExInfo::RetrieveMethList()
{
    CONTRACTL
    {
        THROWS;
        GC_TRIGGERS;
        MODE_COOPERATIVE;
    }
    CONTRACTL_END;

    PTRARRAYREF oMethList;

    // Retrieve the expando OBJECTREF.
    OBJECTREF TargetObj = GetReflectionObject();
    GCPROTECT_BEGIN(TargetObj);

    // Retrieve the GetMembers MethodDesc.
    MethodDesc *pMD = GetIReflectMD(METHOD__IREFLECT__GET_METHODS);
    MethodDescCallSite getMethods(pMD, &TargetObj);

    // Prepare the arguments that will be passed to the method.
    ARG_SLOT Args[] =
    {
        ObjToArgSlot(TargetObj),
        (ARG_SLOT)BINDER_DefaultLookup
    };

    // Retrieve the array of members from the expando object
    oMethList = (PTRARRAYREF) getMethods.Call_RetOBJECTREF(Args);

    GCPROTECT_END();

    return oMethList;
}

// Virtual method to retrieve the InvokeMember method desc.
MethodDesc* DispatchExInfo::GetInvokeMemberMD()
{
    CONTRACT(MethodDesc*)
    {
        THROWS;
        GC_TRIGGERS;
        MODE_ANY;
        POSTCONDITION(CheckPointer(RETVAL));
    }
    CONTRACT_END;

    RETURN GetIReflectMD(METHOD__IREFLECT__INVOKE_MEMBER);
}

// Virtual method to retrieve the object associated with this DispatchInfo that
// implements IReflect.
OBJECTREF DispatchExInfo::GetReflectionObject()
{
    CONTRACTL
    {
        THROWS;
        GC_TRIGGERS;
        MODE_COOPERATIVE;
    }
    CONTRACTL_END;

    // Runtime type is very special. Because of how it is implemented, calling methods
    // through IDispatch on a runtime type object doesn't work like other IReflect implementors
    // work. To be able to invoke methods on the runtime type, we need to invoke them
    // on the runtime type that represents runtime type. This is why for runtime type,
    // we get the exposed class object and not the actual objectred contained in the
    // wrapper.

    if (m_pMT == g_pRuntimeTypeClass)
        return m_pMT->GetManagedClassObject();
    else
        return m_pSimpleWrapperOwner->GetObjectRef();
}

// Virtual method to retrieve the member info map.
ComMTMemberInfoMap *DispatchExInfo::GetMemberInfoMap()
{
    LIMITED_METHOD_CONTRACT;

    // There is no member info map for IExpando objects.
    return NULL;
}
