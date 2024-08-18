using Reactive.Bindings;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Threading;
using System.Threading.Tasks;
using Windows.Foundation.Metadata;
using Windows.UI.Notifications;
using Windows.UI.Notifications.Management;

namespace UserNotificationListenerTest
{
  public class NotificationListener : IDisposable
  {
    private int processId = 0;
    private UserNotificationListener _listener;
    private readonly CancellationTokenSource _cancelToken = new CancellationTokenSource();
    private readonly ReactiveCollection<Notice> _notices
        = new ReactiveCollection<Notice>();

    public ReadOnlyReactiveCollection<Notice> Notices { get; }

    public NotificationListener()
        => Notices = _notices.ToReadOnlyReactiveCollection();

    public async void Listen()
    {
      var cancelToken = _cancelToken.Token;
      var keyWord = $"{ConfigurationManager.AppSettings["KeyWord"]}";

      _listener ??= await GetUserNotificationListener().ConfigureAwait(false);

      while (!cancelToken.IsCancellationRequested)
      {
        var canRestart = false;
        var previousIds = new HashSet<uint>(_notices.Select(n => n.Id));
        var currentIds = new HashSet<uint>();
        var notices = (await _listener.GetNotificationsAsync(NotificationKinds.Toast)).ToList();

        foreach (var notice in notices)
        {
          var id = notice.Id;

          currentIds.Add(id);

          if (previousIds.Contains(id)) continue;

          var toast = notice.Notification.Visual.GetBinding(KnownNotificationBindings.ToastGeneric);

          if (toast != null)
          {
            var elems = toast.GetTextElements();
            var title = elems.FirstOrDefault()?.Text ?? throw new NullReferenceException();
            var body = string.Join('\n', elems.Skip(1).Select(e => e.Text));

            canRestart = title.Contains(keyWord) || body.Contains(keyWord);
            _notices.AddOnScheduler(new Notice(id, title, body));
          }
        }

        previousIds.ExceptWith(currentIds);
        var removedNotices = previousIds
            .Select(id => _notices.FirstOrDefault(n => n.Id == id))
            .Where(notice => notice != null);
        foreach (var removedNotice in removedNotices)
        {
          _notices.RemoveOnScheduler(removedNotice);
        }

        await Task.Delay(1000);

        if (canRestart)
        {
          await Restart();
        }
      }
    }

    private async Task Restart()
    {
      if (processId != 0)
      {
        Process.GetProcessById(processId).Kill();
        processId = 0;
      }

      await Task.Delay(1000);

      var app = new Process();
      var appLocation = $"{ConfigurationManager.AppSettings["AppLocation"]}";

      app.StartInfo.FileName = appLocation;
      app.Start();

      processId = app.Id;
    }

    private async ValueTask<UserNotificationListener> GetUserNotificationListener()
    {
      const string typeName = "Windows.UI.Notifications.Management.UserNotificationListener";
      var isSupported = ApiInformation.IsTypePresent(typeName);
      if (!isSupported) throw new NotSupportedException($"{typeName} is not supported.");

      var listener = UserNotificationListener.Current;

      var accessStatus = await listener.RequestAccessAsync();
      if (accessStatus != UserNotificationListenerAccessStatus.Allowed)
      {
        throw new InvalidOperationException("Access to notification is denied.");
      }

      return listener;
    }

    public void Dispose() => _cancelToken.Cancel();

    public class Notice : INotifyPropertyChanged
    {
      // メモリリーク対策
      public event PropertyChangedEventHandler PropertyChanged;

      public uint Id { get; }
      public string Title { get; }
      public string Body { get; }

      public Notice(uint id, string title, string body)
      {
        Id = id;
        Title = title;
        Body = body;
      }
    }
  }
}
