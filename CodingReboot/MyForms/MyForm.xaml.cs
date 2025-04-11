using System;
using System.Collections.Generic;
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
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Win32;

namespace CodingReboot.MyForms
{
    /// <summary>
    /// Interaction logic for MyForm.xaml
    /// </summary>
    public partial class MyForm : Window
    {
        public MyForm()
        {
            InitializeComponent();
        }

        //step 03: create methods for buttons
        private void btnSelect_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFile = new OpenFileDialog(); //open file browser
            openFile.InitialDirectory = "C:\\";
            //back slash is special character so you have to use two or you can type @ at the beginning
            openFile.Filter = "csv files(*.csv)|*.csv"; //text shown in the text field |* the actual file format you want

            if (openFile.ShowDialog() == true)
            {
                tbxFile.Text = openFile.FileName;
            }
            else
            {
                tbxFile.Text = "No file selected";
            }
        }

        private void btnOK_Click(object sender, RoutedEventArgs e)
        {
            //tell add-in that user clicked ok
            this.DialogResult = true;         
            //form to close
            this.Close();
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            //form to close
            this.DialogResult = false;
            this.Close();
            //tell add-in that user clicked cancel
        }

        public string getTexBoxValue()
        {
            return tbxFile.Text;
        }

        public bool getCheckbox1()
        {
            if (chbCheck1.IsChecked == true) 
                return true;
            else 
                return false;
        }

        public string GetGroup1()
        {
            if (rb1.IsChecked == true)
                return rb1.Content.ToString();
            else if(rb2.IsChecked == true)
                return rb2.Content.ToString();
            else 
                return rb3.Content.ToString();

        }
    }
}
