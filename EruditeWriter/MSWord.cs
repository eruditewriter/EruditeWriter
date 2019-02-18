using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;
using Word = Microsoft.Office.Interop.Word;

namespace EruditeWriter
{
    class MSWordClass
    {
        public Word.Application App        = null;
        public Word.Document    currentDoc = null;
        private     Timer       saveTimer  = null;
        private     Timer       statTimer = null;

        private EWApp thisApp = (EWApp)Application.Current;

        public bool WordInit(bool versiontest)
        {
            Word.Application tempApp = new Word.Application();
            if (versiontest) //we want to test for correct version number
            {
                decimal version;
                //convert the Version from string to decimal as it will be listed as "15.0"
                decimal.TryParse(tempApp.Version, out version);
                if (version < 15) //15 equals Word 2013
                {
                    {
                        //Word 2013 or greater is not installed
                        //release our MSWord objects
                        App.Quit();
                        App = null; //Quit event delegate not setup yet so null App

                        //Show message that correct Word version is not installed
                        MessageBox.Show("You must have Microsoft Word 2013 or greater installed");
                        return false; //failure
                    }
                }
            }
            //Correct Word version is installed.
            App = new Word.Application(); //we have to start two instances to overcome event delegates being general
            //Now set-up delegates for events
            //following best practices we use ((Word.ApplicationEvents4_Event)App) if there is a name conflict
            ((Word.ApplicationEvents4_Event)App).Quit += EventAppQuit;  //handle quitting Word App
            ((Word.ApplicationEvents4_Event)App).NewDocument += EventNewDocument; //handle the user making a new document from within Word
            App.DocumentOpen += EventDocumentOpen; //handle the user opening a new document from within Word
            tempApp.Quit();
            tempApp = null;
            return true; //success
        }//end WordInit

        private void EventDocumentOpen(Word.Document Doc)
        {
            //If the user opens a document from WITHIN Word close it immediately
            //essentially prevent the user from using this functionality
            Doc.Close();
        }//end EventDocumentOpen

        public void EventAppQuit()
        {
            //If the user quits Word (i.e. closes all open documents) set the App variable to null and
            //reinstantiate a new instance of Word
            App = null;
            WordInit(false); //no need to test version number
        }//end EventAppQuit

        public void EventNewDocument(Word.Document Doc)
        {
            //If the user creates a new document from WITHIN Word close it immediately
            //essentially prevent the user from using this functionality
            Doc.Close();
        }//end EventNewDocument

        //This runs when this application is quitting and properly shutsdown Word
        public int WordQuit()
        {
            if (App != null) //check if Word is running
            {
                try
                {
                    App.Quit(SaveChanges: Word.WdSaveOptions.wdPromptToSaveChanges);
                }
                catch (System.Runtime.InteropServices.COMException e)
                {
                    //convert exception to integer and return
                    return e.HResult;
                }
            }
            return 0;
        }//end WordQuit

        //used to handle the opening of documents
        public bool OpenDoc(string filename, bool read_only)
        {
            //check if we have a document already open
            if (currentDoc != null)
            {
                //check if the current open doc is the same file we are being asked to open
                if (String.Equals(currentDoc.FullName, filename, StringComparison.OrdinalIgnoreCase))
                {
                    //file is already open do nothing
                    return false;
                }
                //it's not the same file so close it  (this is the one document open model)
                CloseDoc();
            }
            //now check if the file exists
            if (!File.Exists(filename))
            {
                //file does not exist so create it
                try
                {
                    using (FileStream fs = File.Create(filename))
                    {
                        if (fs == null)
                        {
                            MessageBox.Show("Error Creating Document");
                            return false;
                        }
                    }
                }
                catch (Exception e)
                {
                    MessageBox.Show("Error Creating Document! Error: " + e.ToString());
                    return false;
                }
            }
            App.Visible = true; //make sure that Word is visible
            App.DocumentOpen -= EventDocumentOpen; //clear the document open event so we can open the new document
            if (read_only == false)
            {
                //open the document as read-write
                currentDoc = App.Documents.Open(filename, AddToRecentFiles: false, Revert: false);
            }
            else
            {
                //open the document as read-only
                currentDoc = App.Documents.Open(filename, AddToRecentFiles: false, Revert: false, ReadOnly: true);
            }
            //setup the close delegate to clear the currentDoc variable
            ((Word.DocumentEvents2_Event)currentDoc).Close += delegate { currentDoc = null; saveTimer.Dispose(); statTimer.Dispose(); };
            App.DocumentOpen += EventDocumentOpen; //reset the open event
            saveTimer = new Timer(autoSaveDoc, currentDoc, 10000, 10000);
            statTimer = new Timer(statUpdateTimer, thisApp, 1000, 1000);
            return true;
        }//end NewDoc

        public void CloseDoc()
        {
            currentDoc.Close(SaveChanges: Word.WdSaveOptions.wdSaveChanges);
            System.Windows.Controls.Label charslabel = thisApp.MainWindow.FindName("DocChars") as System.Windows.Controls.Label;
            System.Windows.Controls.Label wordslabel = thisApp.MainWindow.FindName("DocWords") as System.Windows.Controls.Label;
            System.Windows.Controls.Label sentslabel = thisApp.MainWindow.FindName("DocSentences") as System.Windows.Controls.Label;
            charslabel.Content = 0;
            wordslabel.Content = 0;
            sentslabel.Content = 0;
        }//end CloseDoc

        public bool FinishDocs(List<string> filestomerge, string finishname)
        {
            try
            {
                App.DocumentOpen -= EventDocumentOpen;
                Word.Document finishdocument = App.Documents.Open(finishname, AddToRecentFiles: false, Revert: false);
                Word.Selection selection = App.Selection;
                foreach (var file in filestomerge)
                {
                    selection.InsertFile(file);
                }
                finishdocument.Save();
                finishdocument.Close();
                App.DocumentOpen += EventDocumentOpen;
                MessageBox.Show("Finished Document Created");
                OpenDoc(finishname, true);
            }
            catch (Exception e)
            {
                MessageBox.Show("Error promoting. Error:" + e.ToString());
                return false;
            }
            return true;
        }//end FinishDocs

        public bool PromoteDocs(List<string> filestomerge, string folioname)
        {
            try
            {
                if (currentDoc != null)
                {
                    //check if the cuurentdoc is the folio we are merging to
                    if (String.Equals(currentDoc.FullName, folioname, StringComparison.OrdinalIgnoreCase))
                    {
                        //close the folio before modifying
                        CloseDoc();
                    }
                }
                //clear the open event
                App.DocumentOpen -= EventDocumentOpen;
                //now open the folio and loop through the files to insert
                Word.Document foliodoc = App.Documents.Open(folioname, AddToRecentFiles: false, Revert: false);
                //Word.Selection selection = App.Selection;
                int start = foliodoc.Content.End - 1;
                foreach (var file in filestomerge)
                {
                    Word.Range rng = foliodoc.Range(start);
                    rng.InsertParagraphAfter();
                    //selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    if (currentDoc != null)
                    {
                        if (String.Equals(currentDoc.FullName, file, StringComparison.OrdinalIgnoreCase))
                        {
                            //close the doc before modifying
                            CloseDoc();
                        }
                    }
                    rng.InsertFile(file);
                    start = App.ActiveDocument.Content.End - 1;
                }
                foliodoc.Save();
                foliodoc.Close();
                App.DocumentOpen += EventDocumentOpen;
                //open the final document
                OpenDoc(folioname, false);
            }
            catch (Exception e)
            {
                MessageBox.Show("Error promoting. Error:" + e.ToString());
                return false;
            }
            return true;
        }

        public void UpdateStats()
        {
            try
            {
                int allchars, allwords, allsentences;
                //Word.Characters docchars = thisApp.MSWord.currentDoc.Characters;
                //Word.Words docwords = thisApp.MSWord.currentDoc.Words;
                int docchars = thisApp.MSWord.currentDoc.ComputeStatistics(Word.WdStatistic.wdStatisticCharacters, true);
                int docwords = thisApp.MSWord.currentDoc.ComputeStatistics(Word.WdStatistic.wdStatisticWords, true);
                Word.Sentences docsentences = thisApp.MSWord.currentDoc.Sentences;
                thisApp.Codex.GetStatCount(out allchars, out allwords, out allsentences);
                System.Windows.Controls.Label doccharslabel = thisApp.MainWindow.FindName("DocChars") as System.Windows.Controls.Label;
                System.Windows.Controls.Label docwordslabel = thisApp.MainWindow.FindName("DocWords") as System.Windows.Controls.Label;
                System.Windows.Controls.Label docsentslabel = thisApp.MainWindow.FindName("DocSentences") as System.Windows.Controls.Label;
                System.Windows.Controls.Label allcharslabel = thisApp.MainWindow.FindName("AllChars") as System.Windows.Controls.Label;
                System.Windows.Controls.Label allwordslabel = thisApp.MainWindow.FindName("AllWords") as System.Windows.Controls.Label;
                System.Windows.Controls.Label allsentslabel = thisApp.MainWindow.FindName("AllSentences") as System.Windows.Controls.Label;
                doccharslabel.Content = docchars.ToString();
                docwordslabel.Content = docwords.ToString();
                docsentslabel.Content = docsentences.Count.ToString();
                allcharslabel.Content = allchars.ToString();
                allwordslabel.Content = allwords.ToString();
                allsentslabel.Content = allsentences.ToString();
            }
            catch (Exception)
            {
                //do nothing
            }
        }

        private void autoSaveDoc(Object currentDoc)
        {
            try
            {
                if ((currentDoc as Word.Document).Saved)
                {
                    //no changes so no need to update
                    return;
                }
                (currentDoc as Word.Document).Save();
                //if we have a value for openElement update the modified time to reflect the new save time
                if (thisApp.Codex.openElement != null)
                {
                    int docchars = (currentDoc as Word.Document).ComputeStatistics(Word.WdStatistic.wdStatisticCharacters, true);
                    int docwords = (currentDoc as Word.Document).ComputeStatistics(Word.WdStatistic.wdStatisticWords, true);
                    Word.Sentences docsentences = (currentDoc as Word.Document).Sentences;
                    thisApp.Codex.openElement.SetAttributeValue("Chars", docchars.ToString());
                    thisApp.Codex.openElement.SetAttributeValue("Words", docwords.ToString());
                    thisApp.Codex.openElement.SetAttributeValue("Sentences", docsentences.Count.ToString());
                    thisApp.Codex.openElement.SetAttributeValue("Modified", DateTime.Now.ToString(@"MM\/dd\/yyy HH:mm:ss"));
                    //every folio, prelim, or matter has at least two ancestors
                    thisApp.Codex.openElement.Parent.SetAttributeValue("Modified", DateTime.Now.ToString(@"MM\/dd\/yyy HH:mm:ss"));
                    thisApp.Codex.openElement.Parent.Parent.SetAttributeValue("Modified", DateTime.Now.ToString(@"MM\/dd\/yyy HH:mm:ss"));
                    if (String.Equals(thisApp.Codex.openElement.Name.ToString(), "Folio"))
                    {
                        //folio will have three ancestors
                        thisApp.Codex.openElement.Parent.Parent.Parent.SetAttributeValue("Modified", DateTime.Now.ToString(@"MM\/dd\/yyy HH:mm:ss"));
                    }
                    thisApp.Codex.xmlEWC.Save(thisApp.Codex.CodexFileName, System.Xml.Linq.SaveOptions.None);
                }
            }
            catch (Exception)
            {
                //do nothing
            }
        }//end autoSaveDoc

        private void statUpdateTimer(Object ourapp)
        {
            EWApp myapp = (EWApp)ourapp;
            myapp.Dispatcher.Invoke(delegate { myapp.MSWord.UpdateStats(); });
        }
    }
}
