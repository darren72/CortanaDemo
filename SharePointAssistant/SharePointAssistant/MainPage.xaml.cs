using System;
using SharePointAssistant.Office365;
using Windows.UI.Xaml;

// The Blank Page item template is documented at http://go.microsoft.com/fwlink/?LinkId=402352&clcid=0x409

namespace SharePointAssistant
{
    /// <summary>
    /// An empty page that can be used on its own or navigated to within a Frame.
    /// </summary>
    public sealed partial class MainPage
    {
        public MainPage()
        {
            InitializeComponent();

            Instance = this;
        }

        // Quick fix to allow us to access this page from App.Xaml OnActivated() - need to find a better way to do this.
        public static MainPage Instance { get; private set; }

        public void UpdateListView(object source)
        {
            AnnouncementsList.ItemsSource = source;
        }

        public void ClearListView()
        {
            AnnouncementsList.ItemsSource = null;
        }

        private async void GetListItems_OnClick(object sender, RoutedEventArgs e)
        {
            var announcements = await SharePoint.GetListItems("Announcements");
            UpdateListView(announcements);
        }

        private void ClearListItems_OnClickListItems_OnClick(object sender, RoutedEventArgs e)
        {
            ClearListView();
        }
    }
}
