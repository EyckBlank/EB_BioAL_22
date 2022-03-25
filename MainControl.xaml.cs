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

using VMS.TPS;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;
//using System.Windows.Forms;


namespace EB_BioAL_WPF
{
    /// <summary>
    /// Interaction logic for MainControl.xaml
    /// </summary>

    public partial class MainControl : UserControl
    {

        //##################
        public MainControl()
        {
            InitializeComponent();

        }
        //#################

        // Lernmoment
        public void ReportProgress(int newProgress, int maxProgress)
        {
            // myProgress.Maximum = maxProgress;
            // myProgress.Value = newProgress;

            if (this.Dispatcher.CheckAccess())
            {
                myProgress.Maximum = maxProgress;
                myProgress.Value = newProgress;
            }
            else
            {
                this.Dispatcher.Invoke(() => myProgress.Maximum = maxProgress);
                this.Dispatcher.Invoke(() => myProgress.Value = newProgress);
            }

        }

        // Lenmoment
        public void ReportStatus(int newStatus, int maxStatus)
        {
            if (this.Dispatcher.CheckAccess())
            {
                L_Test.Content = newStatus.ToString() + " von " + maxStatus.ToString() + " Schichten";
            }
            else
            {
                this.Dispatcher.Invoke(() => L_Test.Content = newStatus.ToString() + " von " + maxStatus.ToString() + " Schichten");
            }

            var window2 = new Window();
            window2.Show();
            window2.Visibility = Visibility.Collapsed;
            window2.Close();

        }

        //public virtual void OnThresholdReached()
        //{
        //    EventHandler handler = ThresholdReached;
        //    handler?.Invoke(this, e);
        //}

        private void startButton_Click(object sender, RoutedEventArgs e)
        {
            // L_Test.Content = "Läuft";
        }


    }
}