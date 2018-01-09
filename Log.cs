using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.CompilerServices;
using System.Text;
using log4net;
using log4net.Repository.Hierarchy;
using log4net.Core;
using log4net.Appender;
using log4net.Layout;

namespace Hyperfish.ImportExport
{
    public class Log
    {
        public static ILog Logger { get; } = LogManager.GetLogger("ImportExportTool");
    }
}
