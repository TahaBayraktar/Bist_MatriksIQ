using BISTMatriks.Models;

namespace BISTMatriks.Services;

/// <summary>
/// Matriks IQ veri kaynağı arayüzü.
/// Matriks IQ C# SDK'sı elinize geçtiğinde MatriksDataProvider
/// sınıfının içini doldurun — geri kalan hiçbir şeyi değiştirmenize gerek yok.
/// </summary>
public interface IMatriksDataProvider
{
    /// <summary>
    /// Belirtilen hisse için son <paramref name="days"/> günlük
    /// günlük (1D) OHLCV verisini getirir. Kronolojik sırada döner.
    /// </summary>
    Task<List<StockBar>> GetDailyBarsAsync(string ticker, int days = 75);

    /// <summary>
    /// Birden fazla hisseyi toplu çeker (performans optimizasyonu için).
    /// Varsayılan implementasyon GetDailyBarsAsync'i seri çağırır;
    /// SDK paralel destekliyorsa override edin.
    /// </summary>
    Task<Dictionary<string, List<StockBar>>> GetDailyBarsBatchAsync(
        IEnumerable<string> tickers, int days = 75);
}
