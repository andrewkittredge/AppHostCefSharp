using System;
using System.IO;
using System.Reflection;
using System.Windows;
using System.Windows.Forms.Integration;
using ExcelDna.Integration;
using log4net;
using Nancy;
using Nancy.Hosting.Self;
using AppHostCefSharp;
using ExcelInterop = NetOffice.ExcelApi;
using System.ServiceModel;

namespace ExcelDnaExample
{
    public class AddIn : IExcelAddIn
    {
        private static readonly ILog Log = LogManager.
            GetLogger(MethodBase.GetCurrentMethod().DeclaringType);

        private static NancyHost host;
        private static ServiceHost pipeHost;
        public AddIn()
        {
            // Prepare AppData
            if (!Directory.Exists(Settings.AppDataPath))
            {
                Directory.CreateDirectory(Settings.AppDataPath);
            }
            else
            {
                // Delete the browser cache on startup.
                var cachePath = Path.Combine(Settings.AppDataPath, "Cache", "Cache");
                if (!Directory.Exists(cachePath)) return;
                foreach (var filename in Directory.GetFiles(cachePath)) File.Delete(filename);
                Directory.Delete(cachePath);
            }

            // Register event-handlers
            var excel = Excel;

            excel.WorkbookBeforeCloseEvent += (ExcelInterop.Workbook wb, ref bool cancel) =>
            {
                // Shutdown the host when the last workbook is closed.
                // NOTE: This is the only effective way known to shutdown the system (including browser windows) gracefully.
                if (excel.Workbooks.Count == 1 && host != null)
                {
                    StopHost();
                }
            };

            excel.WorkbookOpenEvent += wb =>
            {
                // Ensure the host is (re)started when first workbook is opened.
                if (excel.Workbooks.Count == 1 && host == null)
                {
                    StartHost();
                }
            };

            excel.NewWorkbookEvent += wb =>
            {
                // Ensure the host is (re)started when first workbook is created.
                if (excel.Workbooks.Count == 1 && host == null)
                {
                    StartHost();
                }
            };
        }

        private static string WebHost
            => $"https://localhost:444/query/excelfootnotes";

        internal static ExcelInterop.Application Excel
            => new ExcelInterop.Application(null, ExcelDnaUtil.Application);

        internal static string AssemblyDirectory
        {
            get
            {
                // Nancy assembly is known to be in the program files folder along with the XLL.
                var fullPath = Assembly.GetAssembly(typeof(NancyContext)).Location;
                return Path.GetDirectoryName(fullPath);
            }
        }

        void IExcelAddIn.AutoOpen()
        {
            try
            {
                if (host == null)
                {
                    StartHost();
                }
            }
            catch (Exception ex)
            {
                Log.Error("EXCEPTION", ex);
            }
        }

        void IExcelAddIn.AutoClose()
        {
            if (host != null)
            {
                StopHost();
            }
        }

        public static void ShowExampleForm()
        {
            var geometry = new GeometryPersistence("ExampleWindow", 800, 600);
            var start = $"https://localhost:444/query/excelfootnotes";
            var window = new BrowserWindow(start, geometry, Settings.AppDataFolder)
            {
                Title = "AppHostCefSharp"
            };

            StartNamedPipeHost();
            Show(window);
        }

        private static void Show(Window window)
        {
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                if (Application.Current == null)
                {
                    new Application().ShutdownMode = ShutdownMode.OnExplicitShutdown;
                }

                ElementHost.EnableModelessKeyboardInterop(window);

                if (Application.Current != null)
                {
                    Application.Current.MainWindow = window;
                }

                window.Show();
            });
        }

        private static void StartHost()
        {
            if (host != null)
            {
                throw new Exception("Host already started");
            }
            StaticConfiguration.Caching.EnableRuntimeViewUpdates = true;
            StaticConfiguration.Caching.EnableRuntimeViewDiscovery = true;
            var conf = new HostConfiguration
            {
                UrlReservations =
                {
                    CreateAutomatically = true
                }
            };
            host = new NancyHost(conf, new Uri(WebHost));
            host.Start();
        }

        private static void StopHost()
        {
            if (host == null)
            {
                throw new Exception("Host not running");
            }
            host.Stop();
            host = null;
        }

        private static void StartNamedPipeHost()
        {
            pipeHost = new ServiceHost(
                typeof(StringReverser),
                new Uri[] {
                    new Uri("net.pipe://localhost")
                });
                
            pipeHost.AddServiceEndpoint(typeof(IStringReverser),
                    new NetNamedPipeBinding(),
                    "PipeReverse");
            pipeHost.Open();
        }
    }


    //Following https://web.archive.org/web/20141027055124/http://tech.pro/tutorial/855/wcf-tutorial-basic-interprocess-communication
    [ServiceContract]
    public interface IStringReverser
    {
        [OperationContract]
        string ReverseString(string value);
    }

    public class StringReverser : IStringReverser
    {
        public string ReverseString(string value)
        {
            char[] retVal = value.ToCharArray();
            int idx = 0;
            for (int i = value.Length - 1; i >= 0; i--)
                retVal[idx++] = value[i];

            return new string(retVal);
        }
    }


}
