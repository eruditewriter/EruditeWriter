using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Win32;

namespace EruditeWriter
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class OpeningWindow : Window
    {

        private EWApp thisApp = (EWApp)Application.Current;

        public OpeningWindow()
        {
            InitializeComponent();
        }

        private void btnOpenFile_Click(object sender, RoutedEventArgs e)
        {
            if ((thisApp.dirLocation != "Monograph Directory") && (thisApp.Name != "Monograph Name"))
            {
                if (!(File.Exists(thisApp.dirLocation + thisApp.Name + ".ewm")))
                { //new project create project files
                    FileStream fs = null;
                    try
                    {
                        fs = File.Create(thisApp.dirLocation + thisApp.Name + ".ewm");
                        fs.Close();
                        Directory.CreateDirectory(thisApp.dirLocation + "\\" + thisApp.Name);
                        thisApp.dirLocation += thisApp.Name;
                        Directory.CreateDirectory(thisApp.dirLocation + "\\Draft");
                        fs = File.Create(thisApp.dirLocation + "\\Draft\\" + "1.docx");
                        fs.Close();
                    }
                    catch (NotSupportedException)
                    {
                        MessageBox.Show("Directory refers to a non-file device such as con:, com1:, lpt1:, etc. in a non-NTFS environment.");
                    }
                    catch (PathTooLongException)
                    {
                        MessageBox.Show("The specified path, file name, or both exceeds the system limits!");
                    }
                    catch (IOException)
                    {
                        MessageBox.Show("An I/O error has occurred trying to create the monograph!");
                    }
                    catch (SecurityException)
                    {
                        MessageBox.Show("You do not have the required permission to create this monograph!");
                    }
                    catch (UnauthorizedAccessException)
                    {
                        MessageBox.Show("Access permission is denied!");
                    }
                    finally
                    {
                        if (fs != null)
                        {
                            fs.Close();
                        }
                    }
                    //MainWindow mainWin = new MainWindow();
                    //curApp.mainWin = mainWin;
                    //Application.Current.MainWindow = curApp.mainWin;
                    //thisApp.MSWord = new MSWordApp(thisApp.mainWin.WordHostElement.ActualHeight, thisApp.mainWin.WordHostElement.ActualWidth);
                    thisApp.mainWin.WordHostElement.Child = thisApp.MSWord;

                    this.Close();
                    //curApp.mainWin.Show();
                    //mainWin = null;
                }
                else
                {
                    if (!(Directory.Exists(thisApp.dirLocation + "\\" + thisApp.Name)))
                    {
                        MessageBox.Show("Monograph file exists but not the project directory! Please fix or start a new monograph.");
                    }
                    else
                    {
                        thisApp.dirLocation += thisApp.Name;
                        //Window openwin = Application.Current.MainWindow;
                        //MainWindow mainWin = new MainWindow();
                        //curApp.mainWin = mainWin;
                        //Application.Current.MainWindow = curApp.mainWin;
                        //thisApp.MSWord = new WordApp(thisApp.mainWin.WordHostElement.ActualHeight, thisApp.mainWin.WordHostElement.ActualWidth);

                        thisApp.mainWin.WordHostElement.Child = thisApp.MSWord;

                        this.Close();
                        //curApp.mainWin.Show();
                        //mainWin = null;

                        //this.Close();
                    }
                }
            }
            else
            {
                if ((mLoc_textBox.Text == "Monograph Directory") && (mName_textbox.Text == "Monograph Name"))
                {
                        MessageBox.Show("Please Select a monograph name and a directory.");
                }
                else
                {
                    if ((mName_textbox.Text == "Monograph Name") && (mLoc_textBox.Text != "Monograph Directory"))
                    {
                        MessageBox.Show("Please select a monograph name.");
                    }
                    else
                    {
                        MessageBox.Show("Please select a monograph directory.");
                    }
                }
            }
        }//end btnOpenFile_Click

        private void saveDir_Open(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dlg  = new OpenFileDialog();
            dlg.ValidateNames   = false;
            dlg.CheckFileExists = false;
            dlg.CheckPathExists = true;
            dlg.Filter = "";
            dlg.FileName = "Select Folder or .EWM";
            if (dlg.ShowDialog() == true)
            {
                if (dlg.FileName != "")
                {
                    thisApp.dirLocation = dlg.FileName.Remove((dlg.FileName.LastIndexOf('\\')) + 1);
                    thisApp.Name = dlg.FileName.Replace(thisApp.dirLocation, "");
                    if (thisApp.Name != "Select Folder or .EWM")
                    {
                        if (thisApp.Name.LastIndexOf('.') > -1)
                        {
                            thisApp.Name = thisApp.Name.Remove(thisApp.Name.LastIndexOf('.'));
                        }
                    }
                    else
                    {
                        thisApp.Name = "New_Monograph";
                    }
                    mLoc_textBox.Text = thisApp.dirLocation;
                    mName_textbox.Text = thisApp.Name;
                }
            }
        }//end saveDir_Open

    }//end OpeningWindow
}//end namespace
