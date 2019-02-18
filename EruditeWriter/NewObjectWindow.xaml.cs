using System;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;

namespace EruditeWriter
{
    /// <summary>
    /// Interaction logic for NewMonograph.xaml
    /// </summary>
    public partial class NewObjectWindow : Window
    {
        public NewObjectWindow()
        {
            InitializeComponent();
        }

        private void okButton_Click(object sender, RoutedEventArgs e)
        {
            //when ok is clicked check that we pass the validation rule or display an error message
            ItemName.GetBindingExpression(TextBox.TextProperty).UpdateSource();
            if (!Validation.GetHasError(ItemName))
            {
                this.DialogResult = true;
            }
            else
            {
                MessageBox.Show("Invalid characters in name!");
            }
        }//end okButton_Click
    }//end NewObjectWindow class

    //class to contain the validation rule against illegal filename characters
    public class FileNameRule : ValidationRule
    {
        public override ValidationResult Validate(object value, CultureInfo cultureInfo)
        {
            try
            {
                string fileName = System.IO.Path.GetFileName((string)value);
            }
            catch (ArgumentException e)
            {
                // Path functions will throw this
                // if path contains invalid chars
                return new ValidationResult(false, "Illegal characters or " + e.Message);
            }
            return new ValidationResult(true, null);
        }
    }//end FileNameRule class

    //target to validate rule against
    public class validfilename
    {
        public string filename { get; set; }
    }//end validfilename class
}//end namespace EruditeWriter
