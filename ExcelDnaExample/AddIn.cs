using System;
using System.IO;
using System.Reflection;
using System.Windows;
using System.Windows.Forms.Integration;
using ExcelDna.Integration;
using log4net;

using AppHostCefSharp;
using ExcelInterop = NetOffice.ExcelApi;
using System.ServiceModel;

namespace ExcelDnaExample
{
    public class AddIn : IExcelAddIn
    {
        private static readonly ILog Log = LogManager.
            GetLogger(MethodBase.GetCurrentMethod().DeclaringType);

 
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

        }

        private static string WebHost
            => $"https://localhost:444/query/excelfootnotes";

        internal static ExcelInterop.Application Excel
            => new ExcelInterop.Application(null, ExcelDnaUtil.Application);



        void IExcelAddIn.AutoOpen()
        {

        }

        void IExcelAddIn.AutoClose()
        {

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
