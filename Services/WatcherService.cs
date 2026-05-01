namespace BISTMatriks.Services;

public sealed class WatcherService : IDisposable
{
    private readonly SemaphoreSlim _sem = new(1, 1);
    private FileSystemWatcher? _tarihselW;
    private FileSystemWatcher? _canliW;
    private AppSettings _settings = new();

    public bool IsRunning { get; private set; }

    public void Start(AppSettings settings)
    {
        if (IsRunning) return;
        _settings  = settings;
        IsRunning  = true;

        _ = Task.Run(async () =>
        {
            await UpdateTarihsel();
            await UpdateCanli();
        });

        string dir = AppContext.BaseDirectory;
        if (!Directory.Exists(dir))
            Directory.CreateDirectory(dir);

        _tarihselW = Watch(settings.TarihselCsv, 2000, UpdateTarihsel);
        _canliW    = Watch(settings.CanliCsv,     500,  UpdateCanli);

        Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] İzleme başladı. CSV değişince Excel otomatik güncellenecek.");
    }

    public void Stop()
    {
        _tarihselW?.Dispose();
        _canliW?.Dispose();
        _tarihselW = null;
        _canliW    = null;
        IsRunning  = false;
        Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] İzleme durduruldu.");
    }

    private async Task UpdateTarihsel()
    {
        if (!File.Exists(_settings.TarihselCsv))
        {
            Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] Tarihsel CSV henüz yok — MatriksIQ scriptini çalıştırın: bist_tarihsel.csv");
            return;
        }

        await _sem.WaitAsync();
        try
        {
            Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] Tarihsel CSV değişti — güncelleniyor...");
            var allData = await new CsvDataProvider(_settings.TarihselCsv)
                .GetAllDailyBarsAsync(days: 75);

            using var excel = new ExcelWriter(_settings.WorkbookPath, _settings.WorkbookPath);
            excel.WriteTarihselVeri(allData);
            excel.Save();
            Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] Tarihsel güncellendi — {allData.Count} hisse.");
        }
        catch (Exception ex) { Console.WriteLine($"[HATA] Tarihsel: {ex.Message}"); }
        finally { _sem.Release(); }
    }

    private async Task UpdateCanli()
    {
        if (!File.Exists(_settings.CanliCsv))
        {
            Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] Canlı CSV henüz yok — MatriksIQ scriptini çalıştırın: bist_canli.csv");
            return;
        }

        await _sem.WaitAsync();
        try
        {
            Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] Canlı CSV değişti — güncelleniyor...");

            if (ExcelComWriter.TryWriteCanliVeri(_settings.WorkbookPath, _settings.CanliCsv))
            {
                Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] Canlı güncellendi (COM).");
                return;
            }

            using var excel = new ExcelWriter(_settings.WorkbookPath, _settings.WorkbookPath);
            excel.WriteCanliVeri(_settings.CanliCsv);
            excel.Save();
            Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] Canlı güncellendi.");
        }
        catch (Exception ex) { Console.WriteLine($"[HATA] Canlı: {ex.Message}"); }
        finally { _sem.Release(); }
    }

    private static FileSystemWatcher Watch(string csvPath, int debounceMs, Func<Task> update)
    {
        CancellationTokenSource? cts = null;
        var w = new FileSystemWatcher(
            Path.GetDirectoryName(csvPath)!,
            Path.GetFileName(csvPath)!)
        {
            NotifyFilter        = NotifyFilters.LastWrite | NotifyFilters.Size,
            EnableRaisingEvents = true
        };
        w.Changed += (_, _) =>
        {
            cts?.Cancel();
            cts = new CancellationTokenSource();
            var token = cts.Token;
            Task.Delay(debounceMs, token).ContinueWith(t =>
            {
                if (!t.IsCanceled) _ = update();
            }, TaskScheduler.Default);
        };
        return w;
    }

    public void Dispose() => Stop();
}
