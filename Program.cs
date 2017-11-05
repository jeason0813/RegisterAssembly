// Program.cs
// Main control source codes.
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Configuration;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Collections;
using System.Collections.Specialized;
using System.Reflection;
using System.IO;

namespace RegisterAssembly
{
    // The managed definition of the ICreateTypeLib interface.
    [ComImport()]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("00020406-0000-0000-C000-000000000046")]
    public interface ICreateTypeLib
    {
        IntPtr CreateTypeInfo(string szName, System.Runtime.InteropServices.ComTypes.TYPEKIND tkind);
        void SetName(string szName);
        void SetVersion(short wMajorVerNum, short wMinorVerNum);
        void SetGuid(ref Guid guid);
        void SetDocString(string szDoc);
        void SetHelpFileName(string szHelpFileName);
        void SetHelpContext(int dwHelpContext);
        void SetLcid(int lcid);
        void SetLibFlags(uint uLibFlags);
        void SaveAllChanges();
    }

    partial class Program
    {
        // Command line options.
        const string TYPELIB = "TYPELIB";
        const string CODE_BASE = "CODE_BASE";
        const string VERBOSE = "VERBOSE";
        const string UNREGISTER = "UNREGISTER";

        // Windows API to register a COM type library.
        [DllImport("Oleaut32.dll", CallingConvention = CallingConvention.StdCall, CharSet = CharSet.Unicode)]
        public extern static UInt32 RegisterTypeLib(ITypeLib tlib, string szFullPath, string szHelpDir);

        [DllImport("Oleaut32.dll", CallingConvention = CallingConvention.StdCall, CharSet = CharSet.Unicode)]
        public extern static UInt32 UnRegisterTypeLib(ref Guid libID, UInt16 wVerMajor, UInt16 wVerMinor, int lcid, System.Runtime.InteropServices.ComTypes.SYSKIND syskind);

        public static bool Is32Bits()
        {
            if (IntPtr.Size == 4)
            {
                // 32-bit
                return true;
            }

            return false;
        }

        public static bool Is64Bits()
        {
            if (IntPtr.Size == 8)
            {
                // 64-bit
                return true;
            }

            return false;
        }

        static void DisplayUsage()
        {
            Console.WriteLine("Usage :");
            Console.WriteLine("RegisterAssembly [CODE_BASE] [TYPELIB] [VERBOSE] [UNREGISTER] <path to managed assembly file>");
            Console.WriteLine("where CODE_BASE, TYPELIB, VERBOSE, UNREGISTER are optional parameters");
            Console.WriteLine("Note that the UNREGISTER parameter cannot be used with the CODE_BASE or the CREATE_TYPELIB parameters.");
            Console.WriteLine();
        }

        static bool ProcessArguments(string[] args)
        {
            const int iCountMinimumParametersRequired = 1;

            if (args == null)
            {
                Console.WriteLine(string.Format("Invalid number of parameters."));
                return false;
            }

            if (args.Length < iCountMinimumParametersRequired)
            {
                Console.WriteLine(string.Format("Invalid number of parameters (minimum parameters : [{0:D}]).",
                    iCountMinimumParametersRequired));
                return false;
            }

            for (int i = 0; i < args.Length; i++)
            {
                switch (args[i].Trim())
                {
                    case TYPELIB:
                        {
                            m_bTypeLib = true;
                            break;
                        }

                    case CODE_BASE:
                        {
                            m_bCodeBase = true;
                            break;
                        }

                    case VERBOSE:
                        {
                            m_bVerbose = true;
                            break;
                        }

                    case UNREGISTER:
                        {
                            m_bUnregister = true;
                            break;
                        }

                    default:
                        {
                            if (string.IsNullOrEmpty(m_strTargetAssemblyFilePath))
                            {
                                m_strTargetAssemblyFilePath = args[i].Trim();
                                break;
                            }
                            else
                            {
                                Console.WriteLine(string.Format("Invalid parameter : [{0:S}].", args[i]));
                                return false;
                            }
                        }
                }
            }

            if (m_bUnregister)
            {
                if (m_bCodeBase)
                {
                    Console.WriteLine(string.Format("UNEGISTER flag cannot be used with the CODE_BASE flag."));
                    return false;
                }
            }

            return true;
        }        

        static bool DoWork()
        {
            try
            {
                if (m_bVerbose)
                {
                    Console.WriteLine(string.Format("Target Assembly File : [{0:S}].", m_strTargetAssemblyFilePath));
                }

                if (m_bUnregister)
                {
                    if (PerformAssemblyUnregistration(m_strTargetAssemblyFilePath) == false)
                    {
                        return false;
                    }
                    
                    if (m_bTypeLib == true)
                    {
                        return PerformTypeLibUnRegistration(m_strTargetAssemblyFilePath);
                    }
                    else
                    {
                        return true;
                    }
                }
                else
                {
                    if (PerformAssemblyRegistration(m_strTargetAssemblyFilePath, m_bCodeBase) == false)
                    {
                        return false;
                    }

                    if (m_bTypeLib == true)
                    {
                        return PerformTypeLibCreationAndRegistration(m_strTargetAssemblyFilePath);
                    }
                    else
                    {
                        return true;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(string.Format("An exception occurred. Exception description : [{0:S}].", ex.Message));
                return false;
            }
        }

        static void Main(string[] args)
        {
            if (ProcessArguments(args) == false)
            {
                DisplayUsage();
                return;
            }

            if (ReadConfigSettings() == false)
            {
                return;
            }

            DoWork();
        }

        public static bool m_bVerbose = false;
        private static bool m_bTypeLib = false;
        private static bool m_bCodeBase = false;
        private static bool m_bUnregister = false;
        private static string m_strTargetAssemblyFilePath = null;
        private static ITypeLibExporterNotifySink m_pITypeLibExporterNotifySink = null;
    }
}
