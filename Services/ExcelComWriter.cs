using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;
using System.Linq;

namespace BISTMatriks.Services;

// Çalışan Excel instance'ına COM üzerinden bağlanıp CANLI_VERİ sayfasını günceller.
// Excel açık değilse false döner → çağıran dosya tabanlı yönteme düşer.
public static class ExcelComWriter
{
    private const string SH_CANLI = "📡 CANLI_VERİ";
    private static readonly CultureInfo IC = CultureInfo.InvariantCulture;

    // Sabit yedek eşleme — dinamik okuma başarısız olursa kullanılır
    private static readonly (int ExcelCol, int CsvIdx)[] FallbackColMap =
    {
        (3, 2), (4, 3), (5, 4), (6, 5),
        (7, 6), (8, 7), (9, 8), (10, 9), (11, 10)
    };

    // Marshal.GetActiveObject .NET 10'da yok — P/Invoke ile OLE32'den çekiyoruz
    [DllImport("ole32.dll")]
    private static extern int CLSIDFromProgID(
        [MarshalAs(UnmanagedType.LPWStr)] string lpszProgID, out Guid pclsid);

    [DllImport("oleaut32.dll")]
    private static extern int GetActiveObject(
        ref Guid rclsid, IntPtr pvReserved,
        [MarshalAs(UnmanagedType.IUnknown)] out object ppunk);

    private static object? TryGetActiveObject(string progId)
    {
        if (CLSIDFromProgID(progId, out Guid clsid) != 0) return null;
        return GetActiveObject(ref clsid, IntPtr.Zero, out object obj) == 0 ? obj : null;
    }

    // Türkçe karakter dönüşümü + harf dışı tüm karakterleri sil
    private static string Normalize(string s)
    {
        var up = s.ToUpperInvariant()
                  .Replace("İ", "I").Replace("Ğ", "G").Replace("Ş", "S")
                  .Replace("Ü", "U").Replace("Ö", "O").Replace("Ç", "C");
        return new string(up.Where(c => c >= 'A' && c <= 'Z').ToArray());
    }

    // Excel başlık kısaltmaları → CSV sütun normalized karşılıkları
    private static readonly Dictionary<string, string> ColAliases =
        new(StringComparer.OrdinalIgnoreCase)
    {
        ["ONCKAPANIS"]  = "ONCEKIKAPANIS",   // ÖNC. KAPANIS
        ["AGIRLORT"]    = "AGIRLIKLIORT",    // AĞIRL. ORT
        ["DEGISIM"]     = "DEGISIMYUZDE",    // DEĞİŞİM %
    };

    private static string ApplyAliases(string norm) =>
        ColAliases.TryGetValue(norm, out string? a) ? a : norm;

    // Excel satır 3 başlıkları (object[,]) + CSV başlıkları karşılaştırarak eşleme kurar.
    // dynamic parametre almaz → return tipi statik olarak (int,int)[] kalır.
    private static (int ExcelCol, int CsvIdx)[] BuildColMap(
        object[,] excelHeaders, int firstCol, string[] csvHeader)
    {
        var csvByNorm = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        for (int i = 0; i < csvHeader.Length; i++)
            csvByNorm[Normalize(csvHeader[i].Trim())] = i;

        var map = new List<(int ExcelCol, int CsvIdx)>();
        for (int c = 1; c <= excelHeaders.GetLength(1); c++)
        {
            string? hdr = excelHeaders[1, c]?.ToString();
            if (string.IsNullOrWhiteSpace(hdr)) continue;
            int excelCol = firstCol + c - 1;
            if (excelCol <= 2) continue; // A (SIRA) ve B (HİSSE) atla
            string normHdr = ApplyAliases(Normalize(hdr.Trim()));
            if (csvByNorm.TryGetValue(normHdr, out int csvIdx))
                map.Add((excelCol, csvIdx));
        }
        return map.ToArray();
    }

    public static bool TryWriteCanliVeri(string workbookFullPath, string csvPath)
    {
        dynamic? xlApp = null;
        try
        {
            object? raw = TryGetActiveObject("Excel.Application");
            if (raw == null) return false;
            xlApp = raw;

            // Hedef çalışma kitabını bul
            dynamic? wb = null;
            foreach (dynamic w in xlApp.Workbooks)
            {
                if (string.Equals((string)w.FullName, workbookFullPath,
                        StringComparison.OrdinalIgnoreCase))
                { wb = w; break; }
            }
            if (wb == null) return false;

            // CANLI_VERİ sayfasını bul
            dynamic? ws = null;
            foreach (dynamic s in wb.Sheets)
            {
                if ((string)s.Name == SH_CANLI) { ws = s; break; }
            }
            if (ws == null) return false;

            // CSV oku (FileShare.ReadWrite — MatriksIQ yazarken de okunabilsin)
            string[] allLines = ReadCsvShared(csvPath);
            if (allLines.Length < 2) return false;

            string[] csvHeader = allLines[0].Split(';');

            var csvData = new Dictionary<string, string[]>(StringComparer.OrdinalIgnoreCase);
            foreach (var line in allLines.Skip(1))
            {
                var cols = line.Split(';');
                if (cols.Length < 3) continue;
                string ticker = cols[1].Trim().ToUpper();
                if (!string.IsNullOrEmpty(ticker)) csvData[ticker] = cols;
            }

            // Kullanılan sütun aralığını bul
            int firstCol = (int)ws.UsedRange.Column;
            int colCount = (int)ws.UsedRange.Columns.Count;

            // Excel satır 3 başlıklarını tek COM çağrısıyla oku
            object[,] excelHeaders;
            try
            {
                dynamic hdrRange = ws.Range[
                    ws.Cells[3, firstCol],
                    ws.Cells[3, firstCol + colCount - 1]];
                excelHeaders = (object[,])hdrRange.Value2;
            }
            catch { excelHeaders = new object[1, 0]; }

            // Dinamik sütun haritası; başarısız olursa sabit yedek
            var colMap = BuildColMap(excelHeaders, firstCol, csvHeader);
            if (colMap.Length == 0)
            {
                Console.WriteLine("[COM] UYARI — Dinamik sütun eşleşmesi başarısız, sabit eşleme kullanılıyor.");
                colMap = FallbackColMap;
            }
            else
            {
                Console.WriteLine($"[COM] {colMap.Length} sütun dinamik olarak eşlendi.");
            }

            // B sütunundan ticker → satır haritası (toplu okuma)
            int lastRow = (int)xlApp.WorksheetFunction.CountA(ws.Columns[2]) + 3;
            if (lastRow < 4) lastRow = 4;

            dynamic colBRange = ws.Range[ws.Cells[4, 2], ws.Cells[lastRow, 2]];
            object[,] colBVals = (object[,])colBRange.Value2;

            var tickerRow = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            for (int i = 1; i <= colBVals.GetLength(0); i++)
            {
                string? val = colBVals[i, 1]?.ToString();
                if (!string.IsNullOrEmpty(val))
                    tickerRow[val.Trim()] = i + 3;
            }

            // Her hisse için tüm veri sütunlarını tek range yazısıyla yaz
            int minExcelCol = colMap.Min(c => c.ExcelCol);
            int maxExcelCol = colMap.Max(c => c.ExcelCol);
            int span        = maxExcelCol - minExcelCol + 1;

            // excelCol → colMap içindeki indeks (hangi csvIdx kullanılacak)
            var colMapIdx = new Dictionary<int, int>();
            for (int i = 0; i < colMap.Length; i++)
                colMapIdx[colMap[i].ExcelCol] = i;

            int updated = 0;
            foreach (var (ticker, cols) in csvData)
            {
                if (!tickerRow.TryGetValue(ticker, out int row)) continue;

                var vals = new object[1, span];
                for (int c = 0; c < span; c++)
                {
                    int excelCol = minExcelCol + c;
                    if (!colMapIdx.TryGetValue(excelCol, out int mapI))
                    { vals[0, c] = DBNull.Value; continue; }

                    int csvIdx = colMap[mapI].CsvIdx;
                    if (csvIdx >= cols.Length) { vals[0, c] = DBNull.Value; continue; }

                    string cell = cols[csvIdx].Trim().Replace(',', '.');
                    vals[0, c] = double.TryParse(cell, NumberStyles.Any, IC, out double v)
                        ? (object)v : DBNull.Value;
                }

                dynamic range = ws.Range[
                    ws.Cells[row, minExcelCol],
                    ws.Cells[row, maxExcelCol]];
                range.Value2 = vals;
                updated++;
            }

            Console.WriteLine($"[COM] CANLI_VERİ: {updated} hisse Excel'e yazıldı.");
            return true;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"[COM] Hata: {ex.Message}");
            return false;
        }
        finally
        {
            if (xlApp != null) Marshal.ReleaseComObject(xlApp);
        }
    }

    private static string[] ReadCsvShared(string path)
    {
        using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var sr = new StreamReader(fs, Encoding.UTF8);
        var lines = new List<string>();
        string? line;
        while ((line = sr.ReadLine()) != null) lines.Add(line);
        return lines.ToArray();
    }
}
