using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.VisualBasic.ApplicationServices;

namespace EruditeWriter
{
    // Using VB bits to detect single instances and process accordingly:
    //  * OnStartup is fired when the first instance loads
    //  * OnStartupNextInstance is fired when the application is re-run again
    //    NOTE: it is redirected to this instance thanks to IsSingleInstance
    class SingleInstanceManager : WindowsFormsApplicationBase
    {
        private EWApp _app;

        public SingleInstanceManager()
        {
            IsSingleInstance = true;
        }

        protected override bool OnStartup(StartupEventArgs e)
        {
            // First time app is launched
            _app = new EWApp();
            //set the shutdown mode to OnMainWindowClose which is not the default
            _app.ShutdownMode = System.Windows.ShutdownMode.OnMainWindowClose;
            _app.Run();
            return false;
        }

        protected override void OnStartupNextInstance(StartupNextInstanceEventArgs eventArgs)
        {
            // Subsequent launches
            base.OnStartupNextInstance(eventArgs);
            _app.Activate();
        }
    }
}
