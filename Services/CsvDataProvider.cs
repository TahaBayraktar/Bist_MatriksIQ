using BISTMatriks.Models;

namespace BISTMatriks.Services;

/// <summary>
/// MatriksIQ AlgoTrader'ın ürettiği CSV dosyasını okur.
/// MatriksAlgo_TarihselExport.cs scripti bu formatı üretir.
///
/// CSV format (noktalı virgül ayraç):
///   HISSE;TARIH;ACILIS;YUKSEK;DUSUK;KAPANIS;HACIM;A_ORT
///   GARAN;15.01.2026;47.20;48.10;46.80;47.90;12345678;47.60
/// </summary>
public class CsvDataProvider : IMatriksDataProvider
{
    private readonly string _csvPath;
    private Dictionary<string, List<StockBar>>? _cache;

    public CsvDataProvider(string csvPath)
    {
        _csvPath = csvPath;
    }

    // ── CSV'yi bir kere okuyup önbelleğe al ──────────────────────────────
    private Dictionary<string, List<StockBar>> LoadAll()
    {
        if (_cache != null) return _cache;

        if (!File.Exists(_csvPath))
            throw new FileNotFoundException(
                $"CSV dosyası bulunamadı: {_csvPath}\n" +
                "Önce MatriksAlgo_TarihselExport.cs scriptini MatriksIQ AlgoTrader'da çalıştırın.");

        _cache = new Dictionary<string, List<StockBar>>(StringComparer.OrdinalIgnoreCase);

        int satirNo = 0;
        int hata    = 0;

        using var fs  = new FileStream(_csvPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var rdr = new System.IO.StreamReader(fs, System.Text.Encoding.UTF8);
        foreach (var line in ReadLines(rdr))
        {
            satirNo++;
            if (satirNo == 1) continue; // başlık satırı

            var parts = line.Split(';');
            if (parts.Length < 7) continue;

            try
            {
                string ticker = parts[0].Trim().ToUpper();
                if (string.IsNullOrEmpty(ticker)) continue;

                // Tarih: DD.MM.YYYY
                if (!DateTime.TryParseExact(parts[1].Trim(), "dd.MM.yyyy",
                    System.Globalization.CultureInfo.InvariantCulture,
                    System.Globalization.DateTimeStyles.None, out var date))
                    continue;

                double open  = ParseDouble(parts[2]);
                double high  = ParseDouble(parts[3]);
                double low   = ParseDouble(parts[4]);
                double close = ParseDouble(parts[5]);
                long   vol   = ParseLong(parts[6]);
                double agOrt = parts.Length > 7 ? ParseDouble(parts[7]) : (high + low + close) / 3.0;

                if (close <= 0) continue;

                if (!_cache.TryGetValue(ticker, out var list))
                {
                    list = new List<StockBar>();
                    _cache[ticker] = list;
                }

                list.Add(new StockBar
                {
                    Ticker      = ticker,
                    Date        = date,
                    Open        = open,
                    High        = high,
                    Low         = low,
                    Close       = close,
                    Volume      = vol,
                    WeightedAvg = agOrt
                });
            }
            catch
            {
                hata++;
            }
        }

        // Her hisse için kronolojik sıraya koy
        foreach (var key in _cache.Keys.ToList())
            _cache[key] = _cache[key].OrderBy(b => b.Date).ToList();

        Console.WriteLine($"[CSV] {_csvPath} okundu: {_cache.Count} hisse, {satirNo - 1} satır, {hata} hata");
        return _cache;
    }

    // ── Tüm hisseleri döndür (liste gerekmez) ────────────────────────────
    public Task<Dictionary<string, List<StockBar>>> GetAllDailyBarsAsync(int days = 75)
    {
        var all    = LoadAll();
        var result = new Dictionary<string, List<StockBar>>(StringComparer.OrdinalIgnoreCase);
        foreach (var (ticker, bars) in all)
            result[ticker] = bars.Count > days ? bars.Skip(bars.Count - days).ToList() : bars;
        Console.WriteLine($"[CSV] {result.Count} hisse döndürüldü.");
        return Task.FromResult(result);
    }

    // ── Interface implementasyonu ─────────────────────────────────────────
    public Task<List<StockBar>> GetDailyBarsAsync(string ticker, int days = 75)
    {
        var all = LoadAll();
        if (!all.TryGetValue(ticker, out var bars))
            return Task.FromResult(new List<StockBar>());

        var result = bars.Count > days ? bars.Skip(bars.Count - days).ToList() : bars;
        return Task.FromResult(result);
    }

    public Task<Dictionary<string, List<StockBar>>> GetDailyBarsBatchAsync(
        IEnumerable<string> tickers, int days = 75)
    {
        var all    = LoadAll();
        var result = new Dictionary<string, List<StockBar>>(StringComparer.OrdinalIgnoreCase);

        foreach (var ticker in tickers)
        {
            if (!all.TryGetValue(ticker, out var bars) || bars.Count == 0)
                continue;

            result[ticker] = bars.Count > days ? bars.Skip(bars.Count - days).ToList() : bars;
        }

        Console.WriteLine($"[CSV] {result.Count} hisse döndürüldü.");
        return Task.FromResult(result);
    }

    private static IEnumerable<string> ReadLines(System.IO.StreamReader sr)
    {
        string? line;
        while ((line = sr.ReadLine()) != null) yield return line;
    }

    // ── Yardımcılar ───────────────────────────────────────────────────────
    private static double ParseDouble(string s)
    {
        s = s.Trim().Replace(",", ".");
        return double.TryParse(s, System.Globalization.NumberStyles.Any,
            System.Globalization.CultureInfo.InvariantCulture, out var v) ? v : 0;
    }

    private static long ParseLong(string s)
    {
        s = s.Trim().Replace(".", "").Replace(",", "");
        return long.TryParse(s, out var v) ? v : 0;
    }
}
