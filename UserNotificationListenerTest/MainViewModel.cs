using System.ComponentModel;
using Reactive.Bindings;

namespace UserNotificationListenerTest
{
  public class MainViewModel : INotifyPropertyChanged
  {
    // メモリリーク対策
    public event PropertyChangedEventHandler PropertyChanged;

    private readonly NotificationListener _notification = new NotificationListener();

    public ReadOnlyReactiveCollection<NotificationListener.Notice> Notices
        => _notification.Notices;

    public MainViewModel()
    {
      _notification.Listen();
    }
  }
}
