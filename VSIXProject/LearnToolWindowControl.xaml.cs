using System;
using System.Reflection;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;

namespace LearnVS
{
    public partial class LearnToolWindowControl : UserControl
    {
        public LearnToolWindowControl()
        {
            this.InitializeComponent();
        }

        public Uri BrowserUri
        {
            get => this.LearnBrowser.Source;
            set => this.LearnBrowser.Source = value;
        }

        // awesome code from http://stackoverflow.com/a/28626667 which 
        // fixes internal browser for use in this form
        static LearnToolWindowControl() => SetWebBrowserFeatures();

        static UInt32 GetBrowserEmulationMode()
        {
            int browserVersion = 0;
            using (var ieKey = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Internet Explorer",
                RegistryKeyPermissionCheck.ReadSubTree,
                System.Security.AccessControl.RegistryRights.QueryValues))
            {
                var version = ieKey.GetValue("svcVersion");
                if (null == version)
                {
                    version = ieKey.GetValue("Version");
                    if (null == version)
                        throw new ApplicationException("Microsoft Internet Explorer is required!");
                }
                int.TryParse(version.ToString().Split('.')[0], out browserVersion);
            }

            if (browserVersion < 7)
            {
                throw new ApplicationException("Unsupported version of Microsoft Internet Explorer!");
            }

            UInt32 mode = 11001; // Internet Explorer 11 Edge mode.

            switch (browserVersion)
            {
                case 7:
                    mode = 7000; // Webpages containing standards-based !DOCTYPE directives are displayed in IE7 Standards mode. 
                    break;
                case 8:
                    mode = 8000; // Webpages containing standards-based !DOCTYPE directives are displayed in IE8 mode. 
                    break;
                case 9:
                    mode = 9000; // Internet Explorer 9. Webpages containing standards-based !DOCTYPE directives are displayed in IE9 mode.                    
                    break;
                case 10:
                    mode = 10000; // Internet Explorer 10.
                    break;
            }

            return mode;
        }

        static void SetWebBrowserFeatures()
        {
            // don't change the registry if running in-proc inside Visual Studio
            if (System.ComponentModel.LicenseManager.UsageMode != System.ComponentModel.LicenseUsageMode.Runtime)
                return;

            var appName = System.IO.Path.GetFileName(System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName);

            var featureControlRegKey = @"HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main\FeatureControl\";

            Registry.SetValue(featureControlRegKey + "FEATURE_BROWSER_EMULATION",
                appName, GetBrowserEmulationMode(), RegistryValueKind.DWord);

            // enable the features which are "On" for the full Internet Explorer browser

            Registry.SetValue(featureControlRegKey + "FEATURE_ENABLE_CLIPCHILDREN_OPTIMIZATION",
                appName, 1, RegistryValueKind.DWord);

            Registry.SetValue(featureControlRegKey + "FEATURE_AJAX_CONNECTIONEVENTS",
                appName, 1, RegistryValueKind.DWord);

            Registry.SetValue(featureControlRegKey + "FEATURE_GPU_RENDERING",
                appName, 1, RegistryValueKind.DWord);

            Registry.SetValue(featureControlRegKey + "FEATURE_WEBOC_DOCUMENT_ZOOM",
                appName, 1, RegistryValueKind.DWord);

            Registry.SetValue(featureControlRegKey + "FEATURE_NINPUT_LEGACYMODE",
                appName, 0, RegistryValueKind.DWord);

        }

        private void SetZoom(int zoomLevel)
        {
            const int OLECMDID_OPTICAL_ZOOM = 63;
            const int OLECMDEXECOPT_DONTPROMPTUSER = 2;

            // code for getting IWebBrowser2 interface from https://weblog.west-wind.com/posts/2016/Aug/22/Detecting-and-Setting-Zoom-Level-in-the-WPF-WebBrowser-Control
            var browser = (dynamic)LearnBrowser.GetType().
                GetField("_axIWebBrowser2", BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic).
                GetValue(LearnBrowser);

            // IWebBrowser2::ExecWB documented here: https://msdn.microsoft.com/en-us/library/aa752117(v=vs.85).aspx
            browser.ExecWB(OLECMDID_OPTICAL_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, zoomLevel, ref zoomLevel);
        }

        private void LearnBrowser_Loaded(object sender, RoutedEventArgs e) =>
            LearnBrowser.Source = LearnUris.CSharpTutorial;

        private void LearnBrowser_LoadCompleted(object sender, EventArgs e) => 
            SetZoom(75);
    }
}