using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.Xml;
using System.Xml.Linq;

namespace EruditeWriter
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private EWApp thisApp = (EWApp)Application.Current;

        public MainWindow()
        {
            InitializeComponent();
            //get the persistent window properties
            //we use double.Parse and then catch if the object doesn't exist or the value is corrupt and ignore and go with default
            try
            {
                Top = double.Parse((string)Application.Current.Properties["Top"]);
            }
            catch (ArgumentNullException) { } //property has not been set or has been deleted
            catch (FormatException) { } //not a valid number format
            catch (OverflowException) { } //overflow
            try
            {
               Left = double.Parse((string)Application.Current.Properties["Left"]);
            }
            catch (ArgumentNullException) { }
            catch (FormatException) { }
            catch (OverflowException) { }
            try
            {
                Width = double.Parse((string)Application.Current.Properties["Width"]);
            }
            catch (ArgumentNullException) { }
            catch (FormatException) { }
            catch (OverflowException) { }
            try
            {
                Height = double.Parse((string)Application.Current.Properties["Height"]);
            }
            catch (ArgumentNullException) { }
            catch (FormatException) { }
            catch (OverflowException) { }
        }//end MainWindow()

        private void MenuItem_Close_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }//end MenuItem_Close_Click

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            Application.Current.Properties["Top"]    = this.Top;
            Application.Current.Properties["Left"]   = this.Left;
            Application.Current.Properties["Width"]  = this.Width;
            Application.Current.Properties["Height"] = this.Height;
        }//end Window_Closing

        private void Button_NewCodex_Click(object sender, RoutedEventArgs e)
        {
            // Configure save file dialog box
            Microsoft.Win32.SaveFileDialog dlg = new Microsoft.Win32.SaveFileDialog();
            dlg.CreatePrompt                   = true;
            dlg.DefaultExt                     = ".ewc"; // Default file extension
            dlg.Filter                         = "EruditeWriter Codex |*.ewc"; // Filter files by extension
            dlg.InitialDirectory               = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

            // Show save file dialog box
            bool? result = dlg.ShowDialog();

            // Process save file dialog box results
            if (result == true)
            {
                // Save document
                if (thisApp.Codex.CreateCodex(dlg.FileName) == true)
                {
                    //change our view by "collapsing" and making controls "visible" for working with the Codex
                    New_OpenDock.Visibility  = Visibility.Collapsed;
                    MenuFileStart.Visibility = Visibility.Collapsed;
                    MenuFileCodex.Visibility = Visibility.Visible;
                    MainDock.Visibility      = Visibility.Visible;
                    Col3.Width = GridLength.Auto;
                }
            }
        }//end Button_NewCodex_Click

        private void Button_OpenCodex_Click(object sender, RoutedEventArgs e)
        {
            // Configure open file dialog box
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.DefaultExt                     = ".ewc"; // Default file extension
            dlg.Filter                         = "EruditeWriter Monograph |*.ewc"; // Filter files by extension
            dlg.InitialDirectory               = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

            // Show open file dialog box
            bool? result = dlg.ShowDialog();

            // Process open file dialog box results
            if (result == true)
            {
                thisApp.Codex.OpenCodexXML(dlg.FileName);
            }
            else
            {
                return;
            }
            //change our view by "collapsing" and making controls "visible" for working with the Codex
            New_OpenDock.Visibility  = Visibility.Collapsed;
            MenuFileStart.Visibility = Visibility.Collapsed;
            MenuFileCodex.Visibility = Visibility.Visible;
            MainDock.Visibility      = Visibility.Visible;
            Col3.Width = GridLength.Auto;
        }//end Button_OpenMonograph_Click

        //handle user double-clicking MRU on treeview to open codex .ewc file
        private void MRUFile_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (File.Exists((string)thisApp.Properties[((System.Windows.Controls.TreeViewItem)sender).Name]))
            {
                thisApp.Codex.OpenCodexXML((string)thisApp.Properties[((System.Windows.Controls.TreeViewItem)sender).Name]);
            }
            else
            {
                MessageBox.Show("File no longer exists");
                return;
            }
            //change our view by "collapsing" and making controls "visible" for working with the Codex
            New_OpenDock.Visibility  = Visibility.Collapsed;
            MenuFileStart.Visibility = Visibility.Collapsed;
            MenuFileCodex.Visibility = Visibility.Visible;
            MainDock.Visibility      = Visibility.Visible;
            Col3.Width = GridLength.Auto;
        }//end MRUFile_MouseDoubleClick

        private void CodexList_IsVisibleChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            if ((bool)e.NewValue == true) //CodexList is becoming visible
            {
                int allchars, allwords, allsentences;
                XmlDataProvider xdp = this.Resources["xmlData"] as XmlDataProvider;
                xdp.Source = new Uri(thisApp.Codex.CodexFileName, UriKind.Absolute);
                xdp.XPath = "/EruditeWriterCodex";
                thisApp.Codex.GetStatCount(out allchars, out allwords, out allsentences);
                AllChars.Content = allchars.ToString();
                AllWords.Content = allwords.ToString();
                AllSentences.Content = allsentences.ToString();
            }
        }//end CodexList_IsVisibleChanged

        private void CodexTreeViewRefresh()
        {
            XmlDataProvider xdp = this.Resources["xmlData"] as XmlDataProvider;
            xdp.Refresh();
        }//CodexTreeViewRefresh

        private void MenuItem_NewItem_Click(object sender, RoutedEventArgs e)
        {
            XmlElement item = (sender as MenuItem).DataContext as XmlElement;
            NewObjectWindow dlg = new NewObjectWindow();
            dlg.Owner = this;
            switch(item.Name.ToString())
            {
                case "Monographs":
                    dlg.Title = "New Monograph";
                    dlg.filelabel.Content = "Monograph Name";
                    break;
                case "Monograph":
                    dlg.Title = "New Folio";
                    dlg.filelabel.Content = "Folio Name";
                    break;
                case "FrontMatter":
                    dlg.Title = "New Prelim";
                    dlg.filelabel.Content = "Prelim Name";
                    break;
                case "EndMatter":
                    dlg.Title = "New Matter";
                    dlg.filelabel.Content = "Matter Name";
                    break;
                case "Folio":
                    dlg.Title = "New Section";
                    dlg.filelabel.Content = "Section Name";
                    break;
            }
            dlg.ShowDialog();
            if (dlg.DialogResult == true)
            {
                bool result = false;
                switch (item.Name.ToString())
                {
                    case "Monographs":
                        result = thisApp.Codex.NewMonograph(dlg.ItemName.Text);
                        break;
                    case "Monograph":
                        result = thisApp.Codex.NewFolio(dlg.ItemName.Text, item);
                            break;
                    case "FrontMatter":
                        result = thisApp.Codex.NewPrelim(dlg.ItemName.Text);
                            break;
                    case "EndMatter":
                        result = thisApp.Codex.NewMatter(dlg.ItemName.Text);
                        break;
                    case "Folio":
                        result = thisApp.Codex.NewSection(dlg.ItemName.Text, item);
                        break;
                }
                if (result)
                    CodexTreeViewRefresh();
            }
        }//end MenuItem_NewItem_Click

        private void MenuItem_DeleteItem_Click(object sender, RoutedEventArgs e)
        {
            XmlElement item = (sender as MenuItem).DataContext as XmlElement;
            bool result = false;
            switch (item.Name.ToString())
            {
                case "Monograph":
                    result = thisApp.Codex.DeleteMonograph(item);
                    break;
                case "Section":
                    result = thisApp.Codex.DeleteSection(item);
                    break;
                case "Prelim":
                    result = thisApp.Codex.DeletePrelim(item);
                    break;
                case "Matter":
                    result = thisApp.Codex.DeleteMatter(item); 
                    break;
                case "Folio":
                    result = thisApp.Codex.DeleteFolio(item);
                    break;
            }
            if (result)
                CodexTreeViewRefresh();
        }//end MenuItem_DeleteItem_Click

        private void MenuItem_PublishItem_Click(object sender, RoutedEventArgs e)
        {
            XmlElement item = (sender as MenuItem).DataContext as XmlElement;
            bool result = false;
            switch (item.Name.ToString())
            {
                case "Prelim":
                    result = thisApp.Codex.PublishPrelim(item);
                    break;
                case "Matter":
                    result = thisApp.Codex.PublishMatter(item);
                    break;
                case "Folio":
                    result = thisApp.Codex.PublishFolio(item);
                    break;
            }
            if (result)
                CodexTreeViewRefresh();

        }//end MenuItem_PublishItem_Click

        private void MenuItem_UnPublishItem_Click(object sender, RoutedEventArgs e)
        {
            XmlElement item = (sender as MenuItem).DataContext as XmlElement;
            bool result = false;
            switch (item.Name.ToString())
            {
                case "Prelim":
                    result = thisApp.Codex.UnPublishPrelim(item);
                    break;
                case "Matter":
                    result = thisApp.Codex.UnPublishMatter(item);
                    break;
                case "Folio":
                    result = thisApp.Codex.UnPublishFolio(item);
                    break;
            }
            if (result)
                CodexTreeViewRefresh();
        }//end MenuItem_UnPublishItem_Click

        private void MenuItem_PromoteSection_Click(object sender, RoutedEventArgs e)
        {
            XmlElement item = (sender as MenuItem).DataContext as XmlElement;
            if (thisApp.Codex.PromoteSection(item))
                CodexTreeViewRefresh();
        }//end MenuItem_PromoteSection_Click

        private void MenuItem_CloseCodex_Click(object sender, RoutedEventArgs e)
        {
            thisApp.Codex.xmlEWC.Save(thisApp.Codex.CodexFileName);
            thisApp.MSWord.WordQuit();
            thisApp.Codex.xmlEWC        = null;
            thisApp.Codex.CodexFileName = null;
            thisApp.Codex.dirLocation   = null;
            MenuFileCodex.Visibility    = Visibility.Collapsed;
            MainDock.Visibility         = Visibility.Collapsed;
            New_OpenDock.Visibility     = Visibility.Visible;
            MenuFileCodex.Visibility    = Visibility.Visible;
        }//end MenuItem_CloseCodex_Click

        private void File_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            bool ro = false;
            XmlElement document = (sender as Label).DataContext as XmlElement;
            //get the filename from the xml
            string filename = document.Attributes["FileName"].Value.ToString();
            //check the filename to let us know if it's published and if so set read-only to true
            if (filename.Contains("Published\\Monographs"))
                ro = true;
            thisApp.MSWord.OpenDoc(filename, ro);
            thisApp.Codex.openElement = thisApp.Codex.FindXElement(document);
        }//end FolioFile_MouseDoubleClick

        private void MRUTree_IsVisibleChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            if ((bool)e.NewValue == true) //CodexList is becoming visible
            {
                if (Application.Current.Properties["MRUFile1"] != null)
                {
                    //check that MRUFile exists
                    //strip the path and just set the MRUFile header to the file name
                    this.MRUFile1.Header = ((string)Application.Current.Properties["MRUFile1"]).Replace
                        (((string)Application.Current.Properties["MRUFile1"]).Remove
                            ((((string)Application.Current.Properties["MRUFile1"]).LastIndexOf('\\')) + 1), "");
                    MRUFile1.Visibility = Visibility.Visible;
                }
                if (Application.Current.Properties["MRUFile2"] != null)
                {
                    this.MRUFile2.Header = ((string)Application.Current.Properties["MRUFile2"]).Replace
                        (((string)Application.Current.Properties["MRUFile2"]).Remove
                            ((((string)Application.Current.Properties["MRUFile2"]).LastIndexOf('\\')) + 1), "");
                    MRUFile2.Visibility = Visibility.Visible;
                }
                if (Application.Current.Properties["MRUFile3"] != null)
                {
                    this.MRUFile3.Header = ((string)Application.Current.Properties["MRUFile3"]).Replace
                        (((string)Application.Current.Properties["MRUFile3"]).Remove
                            ((((string)Application.Current.Properties["MRUFile3"]).LastIndexOf('\\')) + 1), "");
                    MRUFile3.Visibility = Visibility.Visible;
                }
                if (Application.Current.Properties["MRUFile4"] != null)
                {
                    this.MRUFile4.Header = ((string)Application.Current.Properties["MRUFile4"]).Replace
                        (((string)Application.Current.Properties["MRUFile4"]).Remove
                            ((((string)Application.Current.Properties["MRUFile4"]).LastIndexOf('\\')) + 1), "");
                    MRUFile4.Visibility = Visibility.Visible;
                }
                if (Application.Current.Properties["MRUFile5"] != null)
                {
                    this.MRUFile5.Header = ((string)Application.Current.Properties["MRUFile5"]).Replace
                        (((string)Application.Current.Properties["MRUFile5"]).Remove
                            ((((string)Application.Current.Properties["MRUFile5"]).LastIndexOf('\\')) + 1), "");
                    MRUFile1.Visibility = Visibility.Visible;
                }
            }
        }//end MRUTree_IsVisibleChanged

        public void MenuItem_Finish_Click(object sender, RoutedEventArgs e)
        {
            thisApp.Codex.FinishCodex();
        }//end MenuItem_Finish_Click

        private void CodexTreeView_ExpandedCollapsed(object sender, RoutedEventArgs e)
        {
            //DataContext gives us the "node" from the tree - check it isn't empty
            if ((e.OriginalSource as TreeViewItem).DataContext != null)
            {
                //convert the source from an XmlElement to an XElement
                XmlElement node = (e.OriginalSource as TreeViewItem).DataContext as XmlElement;
                XElement element = thisApp.Codex.FindXElement(node);
                element.SetAttributeValue("Expanded", (e.OriginalSource as TreeViewItem).IsExpanded);
                thisApp.Codex.xmlEWC.Save(thisApp.Codex.CodexFileName, SaveOptions.None);
            }
        }//end CodexTreeView_ExpandedCollapsed

        private void ButtonPlusMinusItem_Click(object sender, RoutedEventArgs e)
        {
            XmlElement element = (sender as Button).DataContext as XmlElement;
            bool up = true;
            if ((sender as Button).Name == "ButtonMinusItem")
                up = false;
            thisApp.Codex.ElementUpDown(element, up);
            CodexTreeViewRefresh();
        }//end ButtonPlusMinusItem_Click
    }//end MainWindow
}//end namespace EruditeWriter