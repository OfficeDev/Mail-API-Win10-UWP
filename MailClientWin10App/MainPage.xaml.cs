using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using Windows.Foundation;
using Windows.Foundation.Collections;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;
using Windows.UI.Xaml.Controls.Primitives;
using Windows.UI.Xaml.Data;
using Windows.UI.Xaml.Input;
using Windows.UI.Xaml.Media;
using Windows.UI.Xaml.Navigation;

// The Blank Page item template is documented at http://go.microsoft.com/fwlink/?LinkId=402352&clcid=0x409

namespace MailClientWin10App
{
    /// <summary>
    /// An empty page that can be used on its own or navigated to within a Frame.
    /// </summary>
    public sealed partial class MainPage : Page
    {
        public MainPage()
        {
            this.InitializeComponent();
        }

        private void RefreshButton_Click(object sender, RoutedEventArgs e)
        {
            LoadEmailMessagesFromOffice365();
        }
        protected override void OnNavigatedTo(NavigationEventArgs e)
        {
            base.OnNavigatedTo(e);
            LoadEmailMessagesFromOffice365();
        }

        private void MasterListView_ItemClick(object sender, ItemClickEventArgs e)
        {

        }


        private async void LoadEmailMessagesFromOffice365()
        {
            MasterListView.ItemsSource = null;
            progressRing.IsActive = true;
            
            var outlookClient = await AuthUtil.EnsureClient();

            var messages = await outlookClient.Me.Folders["Inbox"].Messages.OrderByDescending(m => m.DateTimeReceived).Take(50).ExecuteAsync();

            progressRing.IsActive = false;

            MasterListView.ItemsSource = messages.CurrentPage;
        }

    }
}
