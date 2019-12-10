using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.ServiceModel;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;

namespace EruditeWriter
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {       
        public App()
        {
            if (SingleInstanceManager.SingleInstance.IsFirstInstance("EWC69674D6-4A98-417B-ADC0-5919F56AE8FE", true))
            {
                SingleInstanceManager.SingleInstance.OnSecondInstanceStarted += NewStartupArgs;

                //start the application
                SplashScreen splashScreen = new SplashScreen("SplashScreen/EW2.png");
                splashScreen.Show(true);
            }
        }

        private void NewStartupArgs(object sender, SingleInstanceManager.SecondInstanceStartedEventArgs e)
        {

        }         
    }
}