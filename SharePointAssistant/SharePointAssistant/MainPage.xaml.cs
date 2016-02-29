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
        }

        private async void GetListItems_OnClick(object sender, RoutedEventArgs e)
        {
            var announcements = await SharePoint.GetAnnouncements();

            AnnouncementsList.ItemsSource = announcements;
        }

        private void ClearListItems_OnClickListItems_OnClick(object sender, RoutedEventArgs e)
        {
            AnnouncementsList.ItemsSource = null;
        }
    }
}
