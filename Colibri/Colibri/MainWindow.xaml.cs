using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace Colibri
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            Driver dr = new Driver();
            if ((!string.IsNullOrEmpty(ResultSheetLoc.Text))&&(!string.IsNullOrEmpty(TransSheetLoc.Text)) &&(!string.IsNullOrEmpty(OrderSheetLoc.Text)))
            {
                if(dr.run(OrderSheetLoc.Text,TransSheetLoc.Text,ResultSheetLoc.Text))
                    System.Windows.Forms.MessageBox.Show("Success");
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("One or more fields missing");
            }
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            using (System.Windows.Forms.OpenFileDialog openFileDialog = new System.Windows.Forms.OpenFileDialog())
            {
                openFileDialog.InitialDirectory = "c:\\";
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    //Get the path of specified file
                    OrderSheetLoc.Text = openFileDialog.FileName;
                }
            }
        }

        private void Button2_Click(object sender, RoutedEventArgs e)
        {
            using (System.Windows.Forms.OpenFileDialog openFileDialog = new System.Windows.Forms.OpenFileDialog())
            {
                openFileDialog.InitialDirectory = "c:\\";
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                { 
                    //Get the path of specified file
                    TransSheetLoc.Text = openFileDialog.FileName;                   
                }
            }
        }

        private void Button3_Click(object sender, RoutedEventArgs e)
        {
            FolderBrowserDialog folderDlg = new FolderBrowserDialog();
            folderDlg.ShowNewFolderButton = true;
            // Show the FolderBrowserDialog.  
            DialogResult result = folderDlg.ShowDialog();
            if (result==System.Windows.Forms.DialogResult.OK)
            {
                ResultSheetLoc.Text = folderDlg.SelectedPath;
            }
        }
    }
}
