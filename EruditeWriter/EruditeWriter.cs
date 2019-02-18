using System;
using System.Collections.Generic;
using System.IO;
using System.IO.IsolatedStorage;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Word = Microsoft.Office.Interop.Word;

namespace EruditeWriter
{
    class EWApp : Application
    {
        //define app wide variables
        public MainWindow  mainWin     = null;     //variable to address the main window
        public MSWordClass MSWord      = null;     //variable to helper class for using Word
        public _Codex      Codex       = null;     //variable to helper class for monographs
        public string      appSettings = "EW.ewa"; //variable pointing to application setting filename

        //startup event
        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);

            //add the function to resolve unhandled exceptions
            Application.Current.DispatcherUnhandledException += App_DispatcherUnhandledException;

            Codex = new _Codex();
            MSWord = new MSWordClass();

            // Restore application-scope property from isolated storage
            IsolatedStorageFile storage = IsolatedStorageFile.GetUserStoreForDomain();
            IsolatedStorageFileStream stream = null;
            try
            {
                stream = new IsolatedStorageFileStream(appSettings, FileMode.Open, storage);
                using (StreamReader reader = new StreamReader(stream))
                {
                    stream = null;
                    // Restore each application-scope property individually
                    while (!reader.EndOfStream)
                    {
                        string[] keyValue = reader.ReadLine().Split(new char[] { ',' });
                        this.Properties[keyValue[0]] = keyValue[1];
                    }
                }
            }
            catch (FileNotFoundException)
            {
                // Handle when file is not found in isolated storage:
                // * When the first application session
                // * When file has been deleted
                storage.CreateFile(appSettings);
            }
            finally
            {
                if (stream != null)
                    stream.Dispose();
            }

            //check to see if Microsoft Word is installed
            Type officeType = Type.GetTypeFromProgID("Word.Application");
            if (officeType == null)
            {
                //Word is not installed
                //Show message that Word is not installed and shutdown
                MessageBox.Show("You must have Microsoft Word 2013 or greater installed");
                Application.Current.Shutdown(-1);
                return;
            }
            else
            {
                //Initialize instance of MS Word and test for version
                if (MSWord.WordInit(true))
                {
                    //show MainWindow and continue with application load
                    mainWin = new MainWindow();
                    mainWin.Show();
                }
            }
            return;
        }//end OnStartup

        protected override void OnExit(ExitEventArgs e)
        {
            base.OnExit(e);

            // Persist application-scope property to isolated storage
            IsolatedStorageFile storage = IsolatedStorageFile.GetUserStoreForDomain();
            IsolatedStorageFileStream stream = null;
            try
            {
                stream = new IsolatedStorageFileStream(appSettings, FileMode.Create, storage);
                using (StreamWriter writer = new StreamWriter(stream))
                {
                    stream = null;
                    // Persist each application-scope property individually
                    foreach (string key in this.Properties.Keys)
                    {
                        writer.WriteLine("{0},{1}", key, this.Properties[key]);
                    }
                }
            }
            finally
            {
                if (stream != null)
                    stream.Dispose();
            }
            //ensure that we quit Word on exit
            ((Word.ApplicationEvents4_Event)MSWord.App).Quit -= MSWord.EventAppQuit;
            int retval = MSWord.WordQuit();
            if (retval != 0) //check if Word has been started by this application and running
            {
                //exception while trying to quit Word
                MessageBox.Show("EruditeWriter is quitting but there was a problem shutting down MS Word");
            }
        }//end OnExit

        private void App_DispatcherUnhandledException(object sender, System.Windows.Threading.DispatcherUnhandledExceptionEventArgs e)
        {
            int retval = MSWord.WordQuit();
            if (retval == -1) //Quitting Word failed
            {
                MessageBox.Show("EruditeWriter encountered an unhandled exception and needs to terminate but there was a problem shutting down MS Word. Exception: " + e.Exception.ToString());
                try
                {
                    Application.Current.Shutdown(-1);
                }
                catch (System.NullReferenceException)
                {
                    //do nothing application is already shutting down
                }
            }
            else
            {
                //exception while trying to quit Word
                MessageBox.Show("EruditeWriter encountered an unhandled exception and will now terminate. Exception: " + e.Exception.ToString());
                e.Handled = true;
                try
                {
                    Application.Current.Shutdown(-1);
                }
                catch (System.NullReferenceException)
                {
                    //do nothing application is already shutting down
                }
            }
            return;
        }//end Application_DispatcherUnhandledException

        public void Activate() //called from SingleInstanceManager
        {
            // Reactivate application's main window
            MainWindow.Activate();
        }//end Activate
    }
}
