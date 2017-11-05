using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.Configuration;
using System.Collections.Specialized;

namespace RegisterAssembly
{
    partial class Program
    {
        const string DEFAULT_TYPELIB_EXPORTER_NOTIFY_SINK = "RegisterAssembly.ConversionEventHandler, RegisterAssembly";

        static bool ReadConfigSettings()
        {
            NameValueCollection appSettings = ConfigurationManager.AppSettings;
            string strTypeLibExporterNotifySink = appSettings["TypeLibExporterNotifySink"];

            if (string.IsNullOrEmpty(strTypeLibExporterNotifySink) == true)
            {
                strTypeLibExporterNotifySink = DEFAULT_TYPELIB_EXPORTER_NOTIFY_SINK;
            }

            Type type = Type.GetType(strTypeLibExporterNotifySink);

            if (type == null)
            {
                return false;
            }

            m_pITypeLibExporterNotifySink = (ITypeLibExporterNotifySink)(Activator.CreateInstance(type, new object[] { m_bVerbose }));

            if (m_pITypeLibExporterNotifySink == null)
            {
                return false;
            }

            return true;
        }
    }
}
