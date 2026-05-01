namespace BISTMatriks.Models;

/// <summary>
/// Bir hissenin tek günlük OHLCV verisi
/// </summary>
public class StockBar
{
    public string Ticker  { get; set; } = "";
    public DateTime Date  { get; set; }
    public double Open    { get; set; }
    public double High    { get; set; }
    public double Low     { get; set; }
    public double Close   { get; set; }
    public long   Volume  { get; set; }
    public double WeightedAvg { get; set; }
}
