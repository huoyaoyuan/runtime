// Licensed to the .NET Foundation under one or more agreements.
// The .NET Foundation licenses this file to you under the MIT license.

using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.Loader.Tracing;

namespace System.Runtime.Loader
{
    internal enum CorPEKind
    {
        peNot = 0x00000000,   // not a PE file
        peILonly = 0x00000001,   // flag IL_ONLY is set in COR header
        pe32BitRequired = 0x00000002,  // flag 32BITREQUIRED is set and 32BITPREFERRED is clear in COR header
        pe32Plus = 0x00000004,   // PE32+ file (64 bit)
        pe32Unmanaged = 0x00000008,    // PE32 without COR header
        pe32BitPreferred = 0x00000010  // flags 32BITREQUIRED and 32BITPREFERRED are set in COR header
    }

    internal struct BundleFileLocation
    {
        public long Size;
        public long Offset;
        public long UncompressedSize;

        public readonly bool IsValid => Offset != 0;
    }

    internal static partial class AssemblyBinderCommon
    {
        [LibraryImport(RuntimeHelpers.QCall, EntryPoint = "PEImage_BinderAcquireImport")]
        internal static unsafe partial IntPtr BinderAcquireImport(IntPtr pPEImage, int* pdwPAFlags);

        [LibraryImport(RuntimeHelpers.QCall, EntryPoint = "PEImage_BinderAcquirePEImage", StringMarshalling = StringMarshalling.Utf16)]
        private static unsafe partial int BinderAcquirePEImage(string szAssemblyPath, out IntPtr ppPEImage, BundleFileLocation bundleFileLocation);

        [LibraryImport(RuntimeHelpers.QCall, EntryPoint = "PEImage_Release")]
        internal static partial void PEImage_Release(IntPtr pPEImage);

        [LibraryImport(RuntimeHelpers.QCall, EntryPoint = "PEImage_GetMVID")]
        internal static partial void PEImage_GetMVID(IntPtr pPEImage, out Guid mvid);

        [LibraryImport(RuntimeHelpers.QCall, EntryPoint = "PEImage_GetPath", StringMarshalling = StringMarshalling.Utf16)]
        internal static unsafe partial char* PEImage_GetPath(IntPtr pPEImage);

        [LibraryImport(RuntimeHelpers.QCall, EntryPoint = "Bundle_AppIsBundle")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static partial bool AppIsBundle();

        [LibraryImport(RuntimeHelpers.QCall, EntryPoint = "Bundle_ProbeAppBundle", StringMarshalling = StringMarshalling.Utf16)]
        private static partial void ProbeAppBundle(string path, [MarshalAs(UnmanagedType.Bool)] bool pathIsBundleRelative, out BundleFileLocation result);

        [LibraryImport(RuntimeHelpers.QCall, EntryPoint = "Bundle_GetAppBundleBasePath")]
        private static partial void GetAppBundleBasePath(StringHandleOnStack path);

        public static bool IsCompatibleAssemblyVersion(BinderAssemblyName requestedName, BinderAssemblyName foundName)
        {
            AssemblyVersion pRequestedVersion = requestedName.Version;
            AssemblyVersion pFoundVersion = foundName.Version;

            if (!pRequestedVersion.HasMajor)
            {
                // An unspecified requested version component matches any value for the same component in the found version,
                // regardless of lesser-order version components
                return true;
            }
            if (!pFoundVersion.HasMajor || pRequestedVersion.Major > pFoundVersion.Major)
            {
                // - A specific requested version component does not match an unspecified value for the same component in
                //   the found version, regardless of lesser-order version components
                // - Or, the requested version is greater than the found version
                return false;
            }
            if (pRequestedVersion.Major < pFoundVersion.Major)
            {
                // The requested version is less than the found version
                return true;
            }

            if (!pRequestedVersion.HasMinor)
            {
                return true;
            }
            if (!pFoundVersion.HasMinor || pRequestedVersion.Minor > pFoundVersion.Minor)
            {
                return false;
            }
            if (pRequestedVersion.Minor < pFoundVersion.Minor)
            {
                return true;
            }

            if (!pRequestedVersion.HasBuild)
            {
                return true;
            }
            if (!pFoundVersion.HasBuild || pRequestedVersion.Build > pFoundVersion.Build)
            {
                return false;
            }
            if (pRequestedVersion.Build < pFoundVersion.Build)
            {
                return true;
            }

            if (!pRequestedVersion.HasRevision)
            {
                return true;
            }
            if (!pFoundVersion.HasRevision || pRequestedVersion.Revision > pFoundVersion.Revision)
            {
                return false;
            }
            return true;
        }

        public static void CreateImageAssembly(IntPtr pPEImage, ref BindResult bindResult) => bindResult.SetResult(new BinderAssembly(pPEImage, isInTPA: false));

        // defined in System.Reflection.PortableExecutable.Machine, but it's in System.Reflection.Metadata
        // also defined in System.Reflection.ImageFileMachine
        private const int IMAGE_FILE_MACHINE_I386 = 0x014c;  // Intel 386.
        private const int IMAGE_FILE_MACHINE_ARMNT = 0x01c4;  // ARM Thumb-2 Little-Endian
        private const int IMAGE_FILE_MACHINE_AMD64 = 0x8664;  // AMD64 (K8)
        private const int IMAGE_FILE_MACHINE_ARM64 = 0xAA64;  // ARM64 Little-Endian

        public static unsafe PEKind TranslatePEToArchitectureType(int* pdwPAFlags)
        {
            CorPEKind CLRPeKind = (CorPEKind)pdwPAFlags[0];
            int dwImageType = pdwPAFlags[1];

            if (CLRPeKind == CorPEKind.peNot)
            {
                // Not a PE. Shouldn't ever get here.
                throw new BadImageFormatException();
            }

            if ((CLRPeKind & CorPEKind.peILonly) != 0 && (CLRPeKind & CorPEKind.pe32Plus) == 0 &&
                (CLRPeKind & CorPEKind.pe32BitRequired) == 0 && dwImageType == IMAGE_FILE_MACHINE_I386)
            {
                // Processor-agnostic (MSIL)
                return PEKind.MSIL;
            }
            else if ((CLRPeKind & CorPEKind.pe32Plus) != 0)
            {
                // 64-bit
                if ((CLRPeKind & CorPEKind.pe32BitRequired) != 0)
                {
                    // Invalid
                    throw new BadImageFormatException();
                }

                // Regardless of whether ILONLY is set or not, the architecture
                // is the machine type.
                if (dwImageType == IMAGE_FILE_MACHINE_ARM64)
                    return PEKind.ARM64;
                else if (dwImageType == IMAGE_FILE_MACHINE_AMD64)
                    return PEKind.AMD64;
                else
                {
                    // We don't support other architectures
                    throw new BadImageFormatException();
                }
            }
            else
            {
                // 32-bit, non-agnostic
                if (dwImageType == IMAGE_FILE_MACHINE_I386)
                    return PEKind.I386;
                else if (dwImageType == IMAGE_FILE_MACHINE_ARMNT)
                    return PEKind.ARM;
                else
                {
                    // Not supported
                    throw new BadImageFormatException();
                }
            }
        }

        // HResult
        public const int RO_E_METADATA_NAME_NOT_FOUND = unchecked((int)0x8000000F);
        public const int E_PATHNOTFOUND = unchecked((int)0x80070003);
        public const int E_NOTREADY = unchecked((int)0x80070015);
        public const int E_BADNETPATH = unchecked((int)0x80070035);
        public const int E_BADNETNAME = unchecked((int)0x80070043);
        public const int E_INVALID_NAME = unchecked((int)0x8007007B);
        public const int E_MODNOTFOUND = unchecked((int)0x8007007E);
        public const int E_DLLNOTFOUND = unchecked((int)0x80070485);
        public const int E_WRONG_TARGET_NAME = unchecked((int)0x80070574);
        public const int INET_E_CANNOT_CONNECT = unchecked((int)0x800C0004);
        public const int INET_E_RESOURCE_NOT_FOUND = unchecked((int)0x800C0005);
        public const int INET_E_OBJECT_NOT_FOUND = unchecked((int)0x800C0006);
        public const int INET_E_DATA_NOT_AVAILABLE = unchecked((int)0x800C0007);
        public const int INET_E_DOWNLOAD_FAILURE = unchecked((int)0x800C0008);
        public const int INET_E_CONNECTION_TIMEOUT = unchecked((int)0x800C000B);
        public const int INET_E_UNKNOWN_PROTOCOL = unchecked((int)0x800C000D);
        public const int FUSION_E_APP_DOMAIN_LOCKED = unchecked((int)0x80131053);
        public const int CLR_E_BIND_ASSEMBLY_VERSION_TOO_LOW = unchecked((int)0x80132000);
        public const int CLR_E_BIND_ASSEMBLY_PUBLIC_KEY_MISMATCH = unchecked((int)0x80132001);
        public const int CLR_E_BIND_ASSEMBLY_NOT_FOUND = unchecked((int)0x80132004);
        public const int CLR_E_BIND_TYPE_NOT_FOUND = unchecked((int)0x80132005);
        public const int CLR_E_BIND_ARCHITECTURE_MISMATCH = unchecked((int)0x80132006);

        public static int BindAssembly(AssemblyLoadContext binder, BinderAssemblyName assemblyName, bool excludeAppPaths, out BinderAssembly? result)
        {
            int kContextVersion = 0;
            BindResult bindResult = default;
            int hr = HResults.S_OK;
            result = null;
            ApplicationContext applicationContext = binder.AppContext;

            // Tracing happens outside the binder lock to avoid calling into managed code within the lock
            using var tracer = new ResolutionAttemptedOperation(assemblyName, binder, ref hr);

        Retry:
            lock (applicationContext.ContextCriticalSection)
            {
                hr = BindByName(applicationContext, assemblyName, false, false, excludeAppPaths, ref bindResult);

                if (hr < 0) goto Exit;

                // Remember the post-bind version
                kContextVersion = applicationContext.Version;
            }

        Exit:
            tracer.TraceBindResult(bindResult);

            if (bindResult.Assembly != null)
            {
                hr = RegisterAndGetHostChosen(applicationContext, kContextVersion, bindResult, out BindResult hostBindResult);

                if (hr == HResults.S_FALSE)
                {
                    // Another bind interfered. We need to retry the entire bind.
                    // This by design loops as long as needed because by construction we eventually
                    // will succeed or fail the bind.
                    bindResult = default;
                    goto Retry;
                }
                else if (hr == HResults.S_OK)
                {
                    Debug.Assert(hostBindResult.Assembly != null);
                    result = hostBindResult.Assembly;
                }
            }

            return hr;
        }

        // Skipped - the managed binder can't bootstrap CoreLib
        // static Assembly? BindToSystem(string systemDirectory);

        // called by vm
        private static unsafe int BindToSystemSatellite(char* systemDirectory, char* simpleName, char* cultureName, out BinderAssembly? assembly)
        {
            // Satellite assembly's relative path

            // append culture name

            // append satellite assembly's simple name

            // append extension
            string relativePath = (string.IsNullOrEmpty(new string(cultureName)) ? new string(simpleName) : new string(cultureName)) + ".dll";

            // Satellite assembly's path:
            //   * Absolute path when looking for a file on disk
            //   * Bundle-relative path when looking within the single-file bundle.
            string sCoreLibSatellite = string.Empty;

            PathSource pathSource = PathSource.Bundle;
            ProbeAppBundle(relativePath, pathIsBundleRelative: true, out BundleFileLocation bundleFileLocation);
            if (!bundleFileLocation.IsValid)
            {
                sCoreLibSatellite = new string(systemDirectory);
                pathSource = PathSource.ApplicationAssemblies;
            }

            sCoreLibSatellite = Path.Combine(sCoreLibSatellite, relativePath);

            int hr = GetAssembly(sCoreLibSatellite, isInTPA: true, out assembly, default);
            if (hr < 0)
            {
                assembly = null;
            }

            AssemblyLoadContext.TracePathProbed(sCoreLibSatellite, pathSource, hr);

            return hr;
        }

        private static int BindByName(
            ApplicationContext applicationContext,
            BinderAssemblyName assemblyName,
            bool skipFailureChecking,
            bool skipVersionCompatibilityCheck,
            bool excludeAppPaths,
            ref BindResult bindResult)
        {
            // Look for already cached binding failure (ignore PA, every PA will lock the context)

            if (applicationContext.FailureCache.TryGetValue((assemblyName.SimpleName, assemblyName.Version), out int hr))
            {
                if (hr < 0) // FAILED(hr)
                {
                    if (hr == HResults.E_FILENOTFOUND && skipFailureChecking)
                    {
                        // Ignore pre-existing transient bind error (re-bind will succeed)
                        applicationContext.FailureCache.Remove((assemblyName.SimpleName, assemblyName.Version));
                    }

                    return hr; // goto LogExit
                }
                else if (hr == HResults.S_FALSE)
                {
                    // workaround: Special case for byte arrays. Rerun the bind to create binding log.
                    assemblyName.IsDefinition = true;
                }
            }

            if (!IsValidArchitecture(assemblyName.ProcessorArchitecture))
            {
                // Assembly reference contains wrong architecture
                hr = HResults.FUSION_E_INVALID_NAME;
                goto Exit;
            }

            hr = BindLocked(applicationContext, assemblyName, skipVersionCompatibilityCheck, excludeAppPaths, ref bindResult);

            if (hr < 0) return hr;

            if (bindResult.Assembly == null)
            {
                // Behavior rules are clueless now
                hr = HResults.E_FILENOTFOUND;
                goto Exit;
            }

        Exit:
            if (hr < 0)
            {
                if (skipFailureChecking)
                {
                    if (hr != HResults.E_FILENOTFOUND)
                    {
                        // Cache non-transient bind error for byte-array
                        hr = HResults.S_FALSE;
                    }
                    else
                    {
                        // Ignore transient bind error (re-bind will succeed)
                        return hr; // goto LogExit;
                    }
                }

                applicationContext.AddToFailureCache(assemblyName, hr);
            }

            return hr;
        }

        private static int BindLocked(
            ApplicationContext applicationContext,
            BinderAssemblyName assemblyName,
            bool skipVersionCompatibilityCheck,
            bool excludeAppPaths,
            ref BindResult bindResult)
        {
            bool isTpaListProvided = applicationContext.TrustedPlatformAssemblyMap != null;
            int hr = FindInExecutionContext(applicationContext, assemblyName, out BinderAssembly? assembly);

            // Add the attempt to the bind result on failure / not found. On success, it will be added after the version check.
            if (hr < 0 || assembly == null)
                bindResult.SetAttemptResult(hr, assembly, isInContext: true);

            if (hr < 0) return hr;

            if (assembly != null)
            {
                if (!skipVersionCompatibilityCheck)
                {
                    // Can't give higher version than already bound
                    bool isCompatible = IsCompatibleAssemblyVersion(assemblyName, assembly.AssemblyName);
                    hr = isCompatible ? HResults.S_OK : FUSION_E_APP_DOMAIN_LOCKED;
                    bindResult.SetAttemptResult(hr, assembly, isInContext: true);

                    // TPA binder returns FUSION_E_REF_DEF_MISMATCH for incompatible version
                    if (hr == FUSION_E_APP_DOMAIN_LOCKED && isTpaListProvided) // hr == FUSION_E_APP_DOMAIN_LOCKED
                        hr = HResults.FUSION_E_REF_DEF_MISMATCH;
                }
                else
                {
                    bindResult.SetAttemptResult(hr, assembly, isInContext: true);
                }

                if (hr < 0) return hr;

                bindResult.SetResult(assembly, isInContext: true);
            }
            else if (isTpaListProvided)
            {
                // BindByTpaList handles setting attempt results on the bind result
                hr = BindByTpaList(applicationContext, assemblyName, excludeAppPaths, ref bindResult);

                if (hr >= 0 && bindResult.Assembly != null) // SUCCEEDED(hr) && pBindResult->HaveResult()
                {
                    bool isCompatible = IsCompatibleAssemblyVersion(assemblyName, bindResult.Assembly.AssemblyName);
                    hr = isCompatible ? HResults.S_OK : FUSION_E_APP_DOMAIN_LOCKED;
                    bindResult.SetAttemptResult(hr, bindResult.Assembly);

                    // TPA binder returns FUSION_E_REF_DEF_MISMATCH for incompatible version
                    if (hr == FUSION_E_APP_DOMAIN_LOCKED && isTpaListProvided) // hr == FUSION_E_APP_DOMAIN_LOCKED
                        hr = HResults.FUSION_E_REF_DEF_MISMATCH;
                }

                if (hr < 0)
                {
                    bindResult.SetNoResult();
                }
            }

            return hr;
        }

        private static int FindInExecutionContext(ApplicationContext applicationContext, BinderAssemblyName assemblyName, out BinderAssembly? assembly)
        {
            applicationContext.ExecutionContext.TryGetValue(assemblyName, out assembly);

            // Set any found context entry. It is up to the caller to check the returned HRESULT
            // for errors due to validation
            if (assembly == null)
                return HResults.S_FALSE;

            if (assembly != null && assemblyName.IsDefinition
                && (assembly.AssemblyName.ProcessorArchitecture != assemblyName.ProcessorArchitecture))
            {
                return FUSION_E_APP_DOMAIN_LOCKED;
            }

            return HResults.S_OK;
        }

        //
        // Tests whether a candidate assembly's name matches the requested.
        // This does not do a version check.  The binder applies version policy
        // further up the stack once it gets a successful bind.
        //
        private static bool TestCandidateRefMatchesDef(BinderAssemblyName requestedAssemblyName, BinderAssemblyName boundAssemblyName, bool tpaListAssembly)
        {
            AssemblyNameIncludeFlags includeFlags = AssemblyNameIncludeFlags.INCLUDE_DEFAULT;

            if (!tpaListAssembly)
            {
                if (requestedAssemblyName.IsNeutralCulture)
                {
                    includeFlags |= AssemblyNameIncludeFlags.EXCLUDE_CULTURE;
                }
            }

            if (requestedAssemblyName.ProcessorArchitecture != PEKind.None)
            {
                includeFlags |= AssemblyNameIncludeFlags.INCLUDE_ARCHITECTURE;
            }

            return boundAssemblyName.Equals(requestedAssemblyName, includeFlags);
        }

        private static int BindSatelliteResourceFromBundle(BinderAssemblyName requestedAssemblyName, string relativePath, ref BindResult bindResult)
        {
            int hr = HResults.S_OK;

            ProbeAppBundle(relativePath, pathIsBundleRelative: true, out BundleFileLocation bundleFileLocation);
            if (!bundleFileLocation.IsValid)
            {
                return hr;
            }

            hr = GetAssembly(relativePath, isInTPA: false, out BinderAssembly? assembly, bundleFileLocation);

            AssemblyLoadContext.TracePathProbed(relativePath, PathSource.Bundle, hr);

            // Missing files are okay and expected when probing
            if (hr == HResults.E_FILENOTFOUND)
            {
                return HResults.S_OK;
            }

            bindResult.SetAttemptResult(hr, assembly);
            if (hr < 0)
                return hr;

            Debug.Assert(assembly != null);
            BinderAssemblyName boundAssemblyName = assembly.AssemblyName;
            if (TestCandidateRefMatchesDef(requestedAssemblyName, boundAssemblyName, tpaListAssembly: false))
            {
                bindResult.SetResult(assembly);
                hr = HResults.S_OK;
            }
            else
            {
                hr = HResults.FUSION_E_REF_DEF_MISMATCH;
            }

            bindResult.SetAttemptResult(hr, assembly);
            return hr;
        }

        private static int BindSatelliteResourceByProbingPaths(
            List<string> resourceRoots,
            BinderAssemblyName requestedAssemblyName,
            string relativePath,
            ref BindResult bindResult,
            PathSource pathSource)
        {
            foreach (string bindingPath in resourceRoots)
            {
                string fileName = Path.Combine(relativePath, bindingPath);
                int hr = GetAssembly(fileName, isInTPA: false, out BinderAssembly? assembly);
                AssemblyLoadContext.TracePathProbed(fileName, pathSource, hr);

                // Missing files are okay and expected when probing
                if (hr == HResults.E_FILENOTFOUND)
                {
                    return HResults.S_OK;
                }

                Debug.Assert(assembly != null);
                BinderAssemblyName boundAssemblyName = assembly.AssemblyName;
                if (TestCandidateRefMatchesDef(requestedAssemblyName, boundAssemblyName, tpaListAssembly: false))
                {
                    bindResult.SetResult(assembly);
                    hr = HResults.S_OK;
                }
                else
                {
                    hr = HResults.FUSION_E_REF_DEF_MISMATCH;
                }

                bindResult.SetAttemptResult(hr, assembly);
                return hr;
            }

            // Up-stack expects S_OK when we don't find any candidate assemblies and no fatal error occurred (ie, no S_FALSE)
            return HResults.S_OK;
        }

        private static int BindSatelliteResource(ApplicationContext applicationContext, BinderAssemblyName requestedAssemblyName, ref BindResult bindResult)
        {
            Debug.Assert(!requestedAssemblyName.IsNeutralCulture);

            string fileName = Path.Combine(requestedAssemblyName.CultureOrLanguage, requestedAssemblyName.SimpleName) + ".dll";

            // Satellite resource probing strategy is to look:
            //   * First within the single-file bundle
            //   * Then under each of the Platform Resource Roots
            //   * Then under each of the App Paths.
            //
            // During each search, if we find a platform resource file with matching file name, but whose ref-def didn't match,
            // fall back to application resource lookup to handle case where a user creates resources with the same
            // names as platform ones.

            int hr = BindSatelliteResourceFromBundle(requestedAssemblyName, fileName, ref bindResult);

            if (bindResult.Assembly != null || hr < 0)
            {
                return hr;
            }

            hr = BindSatelliteResourceByProbingPaths(applicationContext.PlatformResourceRoots, requestedAssemblyName, fileName, ref bindResult, PathSource.PlatformResourceRoots);

            if (bindResult.Assembly != null || hr < 0)
            {
                return hr;
            }

            hr = BindSatelliteResourceByProbingPaths(applicationContext.AppPaths, requestedAssemblyName, fileName, ref bindResult, PathSource.AppPaths);

            return hr;
        }

        private static int BindAssemblyByProbingPaths(List<string> bindingPaths, BinderAssemblyName requestedAssemblyName, out BinderAssembly? result)
        {
            PathSource pathSource = PathSource.AppPaths;

            // Loop through the binding paths looking for a matching assembly
            foreach (string bindingPath in bindingPaths)
            {
                string fileNameWithoutExtension = Path.Combine(bindingPath, requestedAssemblyName.SimpleName);

                // Look for a matching dll first
                string fileName = fileNameWithoutExtension + ".dll";

                int hr = GetAssembly(fileName, isInTPA: false, out BinderAssembly? assembly);
                AssemblyLoadContext.TracePathProbed(fileName, pathSource, hr);

                if (hr < 0)
                {
                    fileName = fileNameWithoutExtension + ".exe";
                    hr = GetAssembly(fileName, isInTPA: false, out assembly);
                    AssemblyLoadContext.TracePathProbed(fileName, pathSource, hr);
                }

                // Since we're probing, file not founds are ok and we should just try another
                // probing path
                if (hr == HResults.COR_E_FILENOTFOUND)
                {
                    continue;
                }

                // Set any found assembly. It is up to the caller to check the returned HRESULT for errors due to validation
                result = assembly;
                if (hr < 0)
                    return hr;

                // We found a candidate.
                //
                // Below this point, we either establish that the ref-def matches, or
                // we fail the bind.

                Debug.Assert(assembly != null);

                // Compare requested AssemblyName with that from the candidate assembly
                if (!TestCandidateRefMatchesDef(requestedAssemblyName, assembly.AssemblyName, tpaListAssembly: false))
                    return HResults.FUSION_E_REF_DEF_MISMATCH;

                return HResults.S_OK;
            }

            result = null;
            return HResults.COR_E_FILENOTFOUND;
        }

        /*
         * BindByTpaList is the entry-point for the custom binding algorithm in CoreCLR.
         *
         * The search for assemblies will proceed in the following order:
         *
         * If this application is a single-file bundle, the meta-data contained in the bundle
         * will be probed to find the requested assembly. If the assembly is not found,
         * The list of platform assemblies (TPAs) are considered next.
         *
         * Platform assemblies are specified as a list of files.  This list is the only set of
         * assemblies that we will load as platform.  They can be specified as IL or NIs.
         *
         * Resources for platform assemblies are located by probing starting at the Platform Resource Roots,
         * a set of folders configured by the host.
         *
         * If a requested assembly identity cannot be found in the TPA list or the resource roots,
         * it is considered an application assembly.  We probe for application assemblies in the
         * AppPaths, a list of paths containing IL files and satellite resource folders.
         *
         */

        public static int BindByTpaList(ApplicationContext applicationContext, BinderAssemblyName requestedAssemblyName, bool excludeAppPaths, ref BindResult bindResult)
        {
            bool fPartialMatchOnTpa = false;

            if (!requestedAssemblyName.IsNeutralCulture)
            {
                int hr = BindSatelliteResource(applicationContext, requestedAssemblyName, ref bindResult);
                if (hr < 0)
                    return hr;
            }
            else
            {
                BinderAssembly? tpaAssembly = null;

                // Is assembly in the bundle?
                // Single-file bundle contents take precedence over TPA.
                // The list of bundled assemblies is contained in the bundle manifest, and NOT in the TPA.
                // Therefore the bundle is first probed using the assembly's simple name.
                // If found, the assembly is loaded from the bundle.
                if (AppIsBundle())
                {
                    // Search Assembly.ni.dll, then Assembly.dll
                    // The Assembly.ni.dll paths are rare, and intended for supporting managed C++ R2R assemblies.
                    ReadOnlySpan<string> candidates = ["ni.dll", ".dll"];

                    // Loop through the binding paths looking for a matching assembly
                    foreach (string candidate in candidates)
                    {
                        string assemblyFileName = requestedAssemblyName.SimpleName + candidate;
                        string? assemblyFilePath = string.Empty;
                        GetAppBundleBasePath(new StringHandleOnStack(ref assemblyFilePath));
                        assemblyFilePath += assemblyFileName;

                        ProbeAppBundle(assemblyFileName, pathIsBundleRelative: true, out BundleFileLocation bundleFileLocation);
                        if (bundleFileLocation.IsValid)
                        {
                            int hr = GetAssembly(assemblyFilePath, isInTPA: true, out tpaAssembly, bundleFileLocation);

                            AssemblyLoadContext.TracePathProbed(assemblyFilePath, PathSource.Bundle, hr);

                            if (hr != HResults.E_FILENOTFOUND)
                            {
                                // Any other error is fatal
                                if (hr < 0) return hr;

                                Debug.Assert(tpaAssembly != null);
                                if (TestCandidateRefMatchesDef(requestedAssemblyName, tpaAssembly.AssemblyName, tpaListAssembly: true))
                                {
                                    // We have found the requested assembly match in the bundle with validation of the full-qualified name.
                                    // Bind to it.
                                    bindResult.SetResult(tpaAssembly);
                                    return HResults.S_OK;
                                }
                            }
                        }
                    }
                }

                // Is assembly on TPA list?
                Debug.Assert(applicationContext.TrustedPlatformAssemblyMap != null);
                if (applicationContext.TrustedPlatformAssemblyMap.TryGetValue(requestedAssemblyName.SimpleName, out TPAEntry tpaEntry))
                {
                    string? tpaFileName = tpaEntry.NIFileName ?? tpaEntry.ILFileName;
                    Debug.Assert(tpaFileName != null);

                    int hr = GetAssembly(tpaFileName, isInTPA: true, out tpaAssembly);
                    AssemblyLoadContext.TracePathProbed(tpaFileName, PathSource.ApplicationAssemblies, hr);

                    bindResult.SetAttemptResult(hr, tpaAssembly);

                    // On file not found, simply fall back to app path probing
                    if (hr != HResults.E_FILENOTFOUND)
                    {
                        // Any other error is fatal
                        if (hr < 0) return hr;

                        Debug.Assert(tpaAssembly != null);
                        if (TestCandidateRefMatchesDef(requestedAssemblyName, tpaAssembly.AssemblyName, tpaListAssembly: true))
                        {
                            // We have found the requested assembly match on TPA with validation of the full-qualified name. Bind to it.
                            bindResult.SetResult(tpaAssembly);
                            bindResult.SetAttemptResult(HResults.S_OK, tpaAssembly);
                            return HResults.S_OK;
                        }
                        else
                        {
                            // We found the assembly on TPA but it didn't match the RequestedAssembly assembly-name. In this case, lets proceed to see if we find the requested
                            // assembly in the App paths.
                            bindResult.SetAttemptResult(HResults.FUSION_E_REF_DEF_MISMATCH, tpaAssembly);
                            fPartialMatchOnTpa = true;
                        }
                    }

                    // We either didn't find a candidate, or the ref-def failed.  Either way; fall back to app path probing.
                }

                if (!excludeAppPaths)
                {
                    // Probe AppPaths

                    int hr = BindAssemblyByProbingPaths(applicationContext.AppPaths, requestedAssemblyName, out BinderAssembly? assembly);
                    bindResult.SetAttemptResult(hr, assembly);

                    if (hr != HResults.E_FILENOTFOUND)
                    {
                        if (hr < 0) return hr;
                        Debug.Assert(assembly != null);

                        // At this point, we have found an assembly with the expected name in the App paths. If this was also found on TPA,
                        // make sure that the app assembly has the same fullname (excluding version) as the TPA version. If it does, then
                        // we should bind to the TPA assembly. If it does not, then bind to the app assembly since it has a different fullname than the
                        // TPA assembly.
                        if (fPartialMatchOnTpa)
                        {
                            Debug.Assert(tpaAssembly != null);

                            if (TestCandidateRefMatchesDef(assembly.AssemblyName, tpaAssembly.AssemblyName, true))
                            {
                                // Fullname (SimpleName+Culture+PKT) matched for TPA and app assembly - so bind to TPA instance.
                                bindResult.SetResult(tpaAssembly);
                                bindResult.SetAttemptResult(hr, tpaAssembly);
                                return HResults.S_OK;
                            }
                            else
                            {
                                // Fullname (SimpleName+Culture+PKT) did not match for TPA and app assembly - so bind to app instance.
                                bindResult.SetResult(assembly);
                                return HResults.S_OK;
                            }
                        }
                        else
                        {
                            // We didn't see this assembly on TPA - so simply bind to the app instance.
                            bindResult.SetResult(assembly);
                            return HResults.S_OK;
                        }
                    }
                }
            }

            // Couldn't find a matching assembly in any of the probing paths
            // Return S_FALSE here. BindByName will interpret a successful HRESULT
            // and lack of BindResult as a failure to find a matching assembly.
            return HResults.S_FALSE;
        }

        private static int GetAssembly(string assemblyPath, bool isInTPA, out BinderAssembly? assembly, BundleFileLocation bundleFileLocation = default)
        {
            int hr = BinderAcquirePEImage(assemblyPath, out IntPtr pPEImage, bundleFileLocation);

            try
            {
                if (hr < 0)
                {
                    // Normalize file not found

                    // ported from Assembly::FileNotFound in coreclr\vm\assembly.cpp
                    if (hr is HResults.E_FILENOTFOUND
                        or E_MODNOTFOUND
                        or E_INVALID_NAME
                        or HResults.CTL_E_FILENOTFOUND
                        or E_PATHNOTFOUND
                        or E_BADNETNAME
                        or E_BADNETPATH
                        or E_NOTREADY
                        or E_WRONG_TARGET_NAME
                        or INET_E_UNKNOWN_PROTOCOL
                        or INET_E_CONNECTION_TIMEOUT
                        or INET_E_CANNOT_CONNECT
                        or INET_E_RESOURCE_NOT_FOUND
                        or INET_E_OBJECT_NOT_FOUND
                        or INET_E_DOWNLOAD_FAILURE
                        or INET_E_DATA_NOT_AVAILABLE
                        or E_DLLNOTFOUND
                        or CLR_E_BIND_ASSEMBLY_VERSION_TOO_LOW
                        or CLR_E_BIND_ASSEMBLY_PUBLIC_KEY_MISMATCH
                        or CLR_E_BIND_ASSEMBLY_NOT_FOUND
                        or RO_E_METADATA_NAME_NOT_FOUND
                        or CLR_E_BIND_TYPE_NOT_FOUND)
                    {
                        hr = HResults.E_FILENOTFOUND;
                    }

                    assembly = null;
                    return hr;
                }

                assembly = new BinderAssembly(pPEImage, isInTPA);
                pPEImage = IntPtr.Zero;
                return HResults.S_OK;
            }
            catch (Exception e)
            {
                assembly = null;
                return e.HResult;
            }
            finally
            {
                // SAFE_RELEASE(pPEImage);
                if (pPEImage != IntPtr.Zero)
                    PEImage_Release(pPEImage);
            }
        }

        public static int Register(ApplicationContext applicationContext, ref BindResult bindResult)
        {
            Debug.Assert(!bindResult.IsContextBound);
            Debug.Assert(bindResult.Assembly != null);

            applicationContext.IncrementVersion();

            // Register the bindResult in the ExecutionContext only if we dont have it already.
            // This method is invoked under a lock (by its caller), so we are thread safe.
            int hr = FindInExecutionContext(applicationContext, bindResult.Assembly.AssemblyName, out BinderAssembly? assembly);
            if (hr < 0)
                return hr;

            if (assembly == null)
            {
                applicationContext.ExecutionContext.Add(bindResult.Assembly.AssemblyName, bindResult.Assembly);
            }
            else
            {
                // Update the BindResult with the assembly we found
                bindResult.SetResult(assembly, isInContext: true);
            }

            return HResults.S_OK;
        }

        public static int RegisterAndGetHostChosen(ApplicationContext applicationContext, int kContextVersion, in BindResult bindResult, out BindResult hostBindResult)
        {
            Debug.Assert(bindResult.Assembly != null);
            hostBindResult = default;
            int hr = HResults.S_OK;

            if (!bindResult.IsContextBound)
            {
                hostBindResult = bindResult;

                // Lock the application context
                lock (applicationContext.ContextCriticalSection)
                {
                    // Only perform costly validation if other binds succeeded before us
                    if (kContextVersion != applicationContext.Version)
                    {
                        hr = OtherBindInterfered(applicationContext, bindResult);
                        if (hr < 0) return hr;

                        if (hr == HResults.S_FALSE)
                        {
                            // Another bind interfered
                            return hr;
                        }
                    }

                    // No bind interfered, we can now register
                    hr = Register(applicationContext, ref hostBindResult);
                    if (hr < 0) return hr;
                }
            }
            else
            {
                // No work required. Return the input
                hostBindResult = bindResult;
            }

            return hr;
        }

        private static int OtherBindInterfered(ApplicationContext applicationContext, BindResult bindResult)
        {
            Debug.Assert(bindResult.Assembly != null);
            Debug.Assert(bindResult.Assembly.AssemblyName != null);

            // Look for already cached binding failure (ignore PA, every PA will lock the context)
            if (!applicationContext.FailureCache.ContainsKey((bindResult.Assembly.AssemblyName.SimpleName, bindResult.Assembly.AssemblyName.Version))) // hr == S_OK
            {
                int hr = FindInExecutionContext(applicationContext, bindResult.Assembly.AssemblyName, out BinderAssembly? assembly);
                if (hr >= 0 && assembly != null)
                {
                    // We can accept this bind in the domain
                    return HResults.S_OK;
                }
            }

            // Some other bind interfered
            return HResults.S_FALSE;
        }

        public static int BindUsingPEImage(AssemblyLoadContext binder, BinderAssemblyName assemblyName, IntPtr pPEImage, bool excludeAppPaths, out BinderAssembly? assembly)
        {
            int hr = HResults.S_OK;

            int kContextVersion = 0;
            BindResult bindResult = default;

            // Prepare binding data
            assembly = null;
            ApplicationContext applicationContext = binder.AppContext;

            // Tracing happens outside the binder lock to avoid calling into managed code within the lock
            using var tracer = new ResolutionAttemptedOperation(assemblyName, binder, ref hr);

            bool mvidMismatch = false;

            try
            {
            Retry:
                mvidMismatch = false;

                // Lock the application context
                lock (applicationContext.ContextCriticalSection)
                {
                    // Attempt uncached bind and register stream if possible
                    // We skip version compatibility check - so assemblies with same simple name will be reported
                    // as a successful bind. Below we compare MVIDs in that case instead (which is a more precise equality check).
                    hr = BindByName(applicationContext, assemblyName, true, true, excludeAppPaths, ref bindResult);

                    if (hr == HResults.E_FILENOTFOUND)
                    {
                        // IF_FAIL_GO(CreateImageAssembly(pPEImage, &bindResult));
                        try
                        {
                            bindResult.SetResult(new BinderAssembly(pPEImage, false));
                        }
                        catch (Exception ex)
                        {
                            return ex.HResult;
                        }
                    }
                    else if (hr == HResults.S_OK)
                    {
                        if (bindResult.Assembly != null)
                        {
                            // Attempt was made to load an assembly that has the same name as a previously loaded one. Since same name
                            // does not imply the same assembly, we will need to check the MVID to confirm it is the same assembly as being
                            // requested.

                            Guid incomingMVID;
                            Guid boundMVID;

                            // GetMVID can throw exception
                            try
                            {
                                PEImage_GetMVID(pPEImage, out incomingMVID);
                                PEImage_GetMVID(bindResult.Assembly.PEImage, out boundMVID);
                            }
                            catch (Exception ex)
                            {
                                return ex.HResult;
                            }

                            mvidMismatch = incomingMVID != boundMVID;
                            if (mvidMismatch)
                            {
                                // MVIDs do not match, so fail the load.
                                return HResults.COR_E_FILELOAD;
                            }

                            // MVIDs match - request came in for the same assembly that was previously loaded.
                            // Let it through...
                        }
                    }

                    // Remember the post-bind version of the context
                    kContextVersion = applicationContext.Version;
                }

                if (bindResult.Assembly != null)
                {
                    // This has to happen outside the binder lock as it can cause new binds
                    hr = RegisterAndGetHostChosen(applicationContext, kContextVersion, bindResult, out BindResult hostBindResult);
                    if (hr < 0) return hr;

                    if (hr == HResults.S_FALSE)
                    {
                        // tracer.TraceBindResult(bindResult);

                        // Another bind interfered. We need to retry entire bind.
                        // This by design loops as long as needed because by construction we eventually
                        // will succeed or fail the bind.
                        bindResult = default;
                        goto Retry;
                    }
                    else if (hr == HResults.S_OK)
                    {
                        assembly = hostBindResult.Assembly;
                    }
                }
            }
            finally
            {
                tracer.TraceBindResult(bindResult, mvidMismatch);
            }

            return hr;
        }

        // CreateDefaultBinder

        public static bool IsValidArchitecture(PEKind architecture)
        {
            if (architecture is PEKind.MSIL or PEKind.None)
                return true;

            PEKind processArchitecture =
#if TARGET_X86
                PEKind.I386;
#elif TARGET_AMD64
                PEKind.AMD64;
#elif TARGET_ARM
                PEKind.ARM;
#elif TARGET_ARM64
                PEKind.ARM64;
#else
                PEKind.MSIL;
#endif

            return architecture == processArchitecture;
        }

        // AssemblySpec::InitializeAssemblyNameRef
        public static unsafe AssemblyName InitializeAssemblyNameRef(BinderAssemblyName assemblyName)
        {
            string culture = string.Empty;

            if ((assemblyName.IdentityFlags & AssemblyIdentityFlags.IDENTITY_FLAG_CULTURE) != 0)
            {
                culture = assemblyName.IsNeutralCulture ? string.Empty : assemblyName.CultureOrLanguage;
            }

            fixed (char* pName = assemblyName.SimpleName)
            fixed (char* pCulture = culture)
            fixed (byte* pPublicKeyOrToken = assemblyName.PublicKeyOrTokenBLOB)
            {
                var nativeAssemblyNameParts = new NativeAssemblyNameParts
                {
                    _pName = pName,
                    _pCultureName = pCulture,
                    _major = (ushort)assemblyName.Version.Major,
                    _minor = (ushort)assemblyName.Version.Minor,
                    _build = (ushort)assemblyName.Version.Build,
                    _revision = (ushort)assemblyName.Version.Revision,
                    _pPublicKeyOrToken = pPublicKeyOrToken,
                    _cbPublicKeyOrToken = assemblyName.PublicKeyOrTokenBLOB.Length,
                };

                if ((assemblyName.IdentityFlags & AssemblyIdentityFlags.IDENTITY_FLAG_PUBLIC_KEY) != 0)
                    nativeAssemblyNameParts._flags |= System.Reflection.AssemblyNameFlags.PublicKey;

                // Architecture unused

                // Retargetable
                if ((assemblyName.IdentityFlags & AssemblyIdentityFlags.IDENTITY_FLAG_RETARGETABLE) != 0)
                    nativeAssemblyNameParts._flags |= System.Reflection.AssemblyNameFlags.Retargetable;

                // Content type unused

                return new AssemblyName(&nativeAssemblyNameParts);
            }
        }
    }

    public partial class AssemblyLoadContext
    {
        [LibraryImport(RuntimeHelpers.QCall, EntryPoint = "DomainAssembly_GetPEAssembly")]
        private static partial IntPtr DomainAssembly_GetPEAssembly(IntPtr pDomainAssembly);

        [LibraryImport(RuntimeHelpers.QCall, EntryPoint = "PEAssembly_GetHostAssembly")]
        private static partial IntPtr PEAssembly_GetHostAssembly(IntPtr pPEAssembly);

        [LibraryImport(RuntimeHelpers.QCall, EntryPoint = "DomainAssembly_EnsureReferenceBinder")]
        private static partial IntPtr DomainAssembly_EnsureReferenceBinder(IntPtr pDomainAssembly, IntPtr pBinder);

        internal static int BindUsingHostAssemblyResolver(
            BinderAssemblyName assemblyName,
            AssemblyLoadContext? defaultBinder,
            AssemblyLoadContext binder,
            out BinderAssembly? loadedAssembly)
        {
            int hr = HResults.E_FAIL;
            loadedAssembly = null;
            BinderAssembly? resolvedAssembly = null;

            // body of RuntimeInvokeHostAssemblyResolver
            bool fResolvedAssembly = false;
            Assembly? refLoadedAssembly = null;
            using var tracer = new ResolutionAttemptedOperation(assemblyName, binder, ref hr);

            // Allocate an AssemblyName managed object
            // Initialize the AssemblyName object
            AssemblyName refAssemblyName = AssemblyBinderCommon.InitializeAssemblyNameRef(assemblyName);

            bool isSatelliteAssemblyRequest = !assemblyName.IsNeutralCulture;
            try
            {
                if (defaultBinder != null)
                {
                    // Step 2 (of CustomAssemblyBinder::BindAssemblyByName) - Invoke Load method
                    // This is not invoked for TPA Binder since it always returns NULL.
                    tracer.GoToStage(Stage.AssemblyLoadContextLoad);

                    refLoadedAssembly = binder.ResolveUsingLoad(refAssemblyName);
                    if (refLoadedAssembly != null)
                    {
                        fResolvedAssembly = true;
                    }

                    hr = fResolvedAssembly ? HResults.S_OK : HResults.COR_E_FILENOTFOUND;

                    // Step 3 (of CustomAssemblyBinder::BindAssemblyByName)
                    if (!fResolvedAssembly && !isSatelliteAssemblyRequest)
                    {
                        tracer.GoToStage(Stage.DefaultAssemblyLoadContextFallback);

                        // If we could not resolve the assembly using Load method, then attempt fallback with TPA Binder.
                        // Since TPA binder cannot fallback to itself, this fallback does not happen for binds within TPA binder.

                        hr = defaultBinder.BindUsingAssemblyName(assemblyName, out BinderAssembly? coreCLRFoundAssembly);
                        if (hr >= 0)
                        {
                            Debug.Assert(coreCLRFoundAssembly != null);
                            resolvedAssembly = coreCLRFoundAssembly;
                            fResolvedAssembly = true;
                        }
                    }
                }

                if (!fResolvedAssembly && isSatelliteAssemblyRequest)
                {
                    // Step 4 (of CustomAssemblyBinder::BindAssemblyByName)

                    // Attempt to resolve it using the ResolveSatelliteAssembly method.
                    tracer.GoToStage(Stage.ResolveSatelliteAssembly);

                    refLoadedAssembly = binder.ResolveSatelliteAssembly(refAssemblyName);
                    if (refLoadedAssembly != null)
                    {
                        // Set the flag indicating we found the assembly
                        fResolvedAssembly = true;
                    }

                    hr = fResolvedAssembly ? HResults.S_OK : HResults.COR_E_FILENOTFOUND;
                }

                if (!fResolvedAssembly)
                {
                    // Step 5 (of CustomAssemblyBinder::BindAssemblyByName)

                    // If we couldn't resolve the assembly using TPA LoadContext as well, then
                    // attempt to resolve it using the Resolving event.
                    tracer.GoToStage(Stage.AssemblyLoadContextResolvingEvent);

                    refLoadedAssembly = binder.ResolveUsingEvent(refAssemblyName);
                    if (refLoadedAssembly != null)
                    {
                        // Set the flag indicating we found the assembly
                        fResolvedAssembly = true;
                    }

                    hr = fResolvedAssembly ? HResults.S_OK : HResults.COR_E_FILENOTFOUND;
                }

                if (fResolvedAssembly && resolvedAssembly == null)
                {
                    // If we are here, assembly was successfully resolved via Load or Resolving events.
                    Debug.Assert(refLoadedAssembly != null);

                    // We were able to get the assembly loaded. Now, get its name since the host could have
                    // performed the resolution using an assembly with different name.

                    RuntimeAssembly? rtAssembly =
                        GetRuntimeAssembly(refLoadedAssembly)
                        ?? throw new InvalidOperationException(SR.Arg_MustBeRuntimeAssembly);

                    IntPtr pDomainAssembly = rtAssembly.GetUnderlyingNativeHandle();
                    IntPtr pLoadedPEAssembly = IntPtr.Zero;
                    bool fFailLoad = false;
                    if (pDomainAssembly == IntPtr.Zero)
                    {
                        // Reflection emitted assemblies will not have a domain assembly.
                        fFailLoad = true;
                    }
                    else
                    {
                        pLoadedPEAssembly = DomainAssembly_GetPEAssembly(pDomainAssembly);
                        if (PEAssembly_GetHostAssembly(pLoadedPEAssembly) == IntPtr.Zero)
                        {
                            // Reflection emitted assemblies will not have a domain assembly.
                            fFailLoad = true;
                        }
                    }

                    // The loaded assembly's BINDER_SPACE::Assembly* is saved as HostAssembly in PEAssembly
                    if (fFailLoad)
                    {
                        // string name = assemblyName.GetDisplayName(AssemblyNameIncludeFlags.INCLUDE_ALL);
                        throw new InvalidOperationException(SR.Host_AssemblyResolver_DynamicallyEmittedAssemblies_Unsupported);
                    }

                    // For collectible assemblies, ensure that the parent loader allocator keeps the assembly's loader allocator
                    // alive for all its lifetime.
                    if (rtAssembly.IsCollectible)
                    {
                        DomainAssembly_EnsureReferenceBinder(pDomainAssembly, binder._nativeAssemblyLoadContext);
                    }

                    resolvedAssembly = GCHandle.FromIntPtr(PEAssembly_GetHostAssembly(pLoadedPEAssembly)).Target as BinderAssembly;
                }

                if (fResolvedAssembly)
                {
                    Debug.Assert(resolvedAssembly != null);

                    // Get the BINDER_SPACE::Assembly reference to return back to.
                    loadedAssembly = resolvedAssembly;
                    hr = HResults.S_OK;

                    tracer.SetFoundAssembly(resolvedAssembly);
                }
                else
                {
                    hr = HResults.COR_E_FILENOTFOUND;
                }
            }
            catch (Exception ex)
            {
                tracer.SetException(ex);
                throw;
            }

            return hr;
        }
    }
}