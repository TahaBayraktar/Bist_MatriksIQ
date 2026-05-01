using BISTMatriks.Models;
using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;

namespace BISTMatriks.Services;

// ── Excel Yazıcı ────────────────────────────────────────────────────────────
// EPPlus KULLANMAZ — saf ZIP + OOXML.
// Çıktı dosyası varsa üzerine yazar (incremental); yoksa şablondan oluşturur.
public class ExcelWriter : IDisposable
{
    private readonly string _templatePath; // ilk oluşturma için kaynak
    private readonly string _outputPath;   // hedef dosya
    private readonly string _sourcePath;   // okuma kaynağı: output(varsa) veya template
    private readonly Dictionary<string, string> _sheetXmlMap;
    private readonly Dictionary<string, byte[]> _pending = new(StringComparer.OrdinalIgnoreCase);

    private const string SH_TARIHSEL = "📊 TARİHSEL_VERİ";
    private const string SH_CANLI    = "📡 CANLI_VERİ";

    private static readonly CultureInfo IC = CultureInfo.InvariantCulture;

    public ExcelWriter(string templatePath, string outputPath)
    {
        _templatePath = templatePath;
        _outputPath   = outputPath;

        // Çıktı dosyası varsa onu kaynak olarak kullan (incremental mod).
        // Excel açıkken dosya kilitli olabilir — IOException'da şablona düş.
        if (File.Exists(outputPath))
        {
            try
            {
                _sheetXmlMap = ReadSheetXmlPaths(outputPath);
                _sourcePath  = outputPath;
            }
            catch (IOException)
            {
                Console.WriteLine("[Excel] Çıktı dosyası kilitli — şablon kullanılıyor (kaynak olarak).");
                _sourcePath  = templatePath;
                _sheetXmlMap = ReadSheetXmlPaths(templatePath);
            }
        }
        else
        {
            _sourcePath  = templatePath;
            _sheetXmlMap = ReadSheetXmlPaths(templatePath);
        }

        var found = new[] { SH_TARIHSEL, SH_CANLI }
            .Where(_sheetXmlMap.ContainsKey).ToList();
        Console.WriteLine($"[Excel] Sayfalar: {string.Join(", ", found)}");

        var missing = new[] { SH_TARIHSEL, SH_CANLI }
            .Where(s => !_sheetXmlMap.ContainsKey(s)).ToList();
        if (missing.Any())
            Console.WriteLine($"[Excel] UYARI — Bulunamadı: {string.Join(", ", missing)}");
    }

    public void Dispose() { }

    // ── CANLI_VERİ: Her zaman şablondan okunur (renkler/stiller korunur).
    // DDE formülleri (#BAŞV!) kaldırılır, CSV değerleri statik olarak yazılır.
    public void WriteCanliVeri(string csvPath)
    {
        if (!_sheetXmlMap.TryGetValue(SH_CANLI, out var xmlPath)) return;

        string xml = ReadEntryFromTemplate(xmlPath);

        // CSV oku (FileShare.ReadWrite — MatriksIQ eş zamanlı yazarken de okunabilsin)
        string[] allLines = ReadCsvShared(csvPath);
        if (allLines.Length < 2) return;
        string[] csvHeader = allLines[0].Split(';');

        var csvData = new Dictionary<string, string[]>(StringComparer.OrdinalIgnoreCase);
        foreach (var line in allLines.Skip(1))
        {
            var cols = line.Split(';');
            if (cols.Length < 3) continue;
            string ticker = cols[1].Trim().ToUpper();
            if (!string.IsNullOrEmpty(ticker)) csvData[ticker] = cols;
        }

        // B sütunundan ticker → satır numarası haritası (inlineStr + sharedStr)
        string[] sharedForTicker = LoadSharedStrings(_sourcePath);
        var tickerRow = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

        foreach (Match m in Regex.Matches(xml,
            @"<c r=""B(\d+)""[^>]*>([\s\S]*?)</c>"))
        {
            if (!int.TryParse(m.Groups[1].Value, out int r)) continue;
            string inner = m.Groups[2].Value;

            string? ticker = null;
            // inlineStr: <is><t>GARAN</t></is>
            var inM = Regex.Match(inner, @"<is><t>([^<]+)</t></is>");
            if (inM.Success)
            {
                ticker = inM.Groups[1].Value.Trim();
            }
            else
            {
                // sharedStr: <v>42</v>
                var vM = Regex.Match(inner, @"<v>(\d+)</v>");
                if (vM.Success && int.TryParse(vM.Groups[1].Value, out int si) && si < sharedForTicker.Length)
                    ticker = sharedForTicker[si].Trim();
            }

            if (!string.IsNullOrEmpty(ticker))
                tickerRow[ticker] = r;
        }

        // Dinamik sütun eşlemesi; başarısız olursa sabit yedek
        var colIdx = BuildColMapFromXml(xml, csvHeader, _templatePath);
        if (colIdx.Length == 0)
        {
            Console.WriteLine("[Excel] UYARI — Dinamik sütun eşleşmesi başarısız, sabit eşleme kullanılıyor.");
            colIdx = new (char Col, int Idx)[]
            {
                ('C', 2), ('D', 3), ('E', 4), ('F', 5),
                ('G', 6), ('H', 7), ('I', 8), ('J', 9), ('K', 10)
            };
        }
        else
        {
            Console.WriteLine($"[Excel] {colIdx.Length} sütun dinamik olarak eşlendi.");
        }

        int updated = 0;
        foreach (var (ticker, cols) in csvData)
        {
            if (!tickerRow.TryGetValue(ticker, out int row)) continue;

            foreach (var (col, idx) in colIdx)
            {
                if (idx >= cols.Length) continue;
                string raw = cols[idx].Trim().Replace(',', '.');
                if (!double.TryParse(raw, NumberStyles.Any, IC, out double val)) continue;
                string valStr  = val.ToString("R", IC);
                string cellRef = $"{col}{row}";

                xml = Regex.Replace(xml,
                    $@"<c r=""{cellRef}""([^>]*)>(?:<f>[^<]*</f>)?<v>[^<]*</v></c>",
                    m =>
                    {
                        var sm = Regex.Match(m.Groups[1].Value, @"s=""(\d+)""");
                        string style = sm.Success ? $" s=\"{sm.Groups[1].Value}\"" : "";
                        return $"<c r=\"{cellRef}\"{style} t=\"n\"><v>{valStr}</v></c>";
                    });
            }
            updated++;
        }

        _pending[xmlPath] = Encoding.UTF8.GetBytes(xml);
        Console.WriteLine($"[Excel] CANLI_VERİ: {updated} hisse güncellendi.");
    }

    // Şablon XML'inin 3. satır başlıklarını ve CSV başlıklarını karşılaştırarak
    // (ExcelSütunHarfi, CsvSütunİndeksi) eşlemesini dinamik olarak kurar.
    private static (char Col, int Idx)[] BuildColMapFromXml(
        string sheetXml, string[] csvHeader, string templateZipPath)
    {
        // CSV: normalize başlık → indeks
        var csvByNorm = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        for (int i = 0; i < csvHeader.Length; i++)
            csvByNorm[Normalize(csvHeader[i].Trim())] = i;

        // Şablon shared strings (başlık hücreleri çoğunlukla shared string olarak saklanır)
        string[] shared = LoadSharedStrings(templateZipPath);

        // Satır 3'teki tüm hücreleri bul
        var rowMatch = Regex.Match(sheetXml, @"<row r=""3""[^>]*>([\s\S]*?)</row>");
        if (!rowMatch.Success) return Array.Empty<(char, int)>();

        var map = new List<(char Col, int Idx)>();

        foreach (Match cm in Regex.Matches(rowMatch.Groups[1].Value,
            @"<c r=""([A-Z]+)3""([^>]*)>([\s\S]*?)</c>"))
        {
            string colRef = cm.Groups[1].Value;
            if (colRef is "A" or "B") continue; // ticker/sıra sütunlarını atla

            string attrs = cm.Groups[2].Value;
            string inner = cm.Groups[3].Value;
            string? text = null;

            if (attrs.Contains("t=\"inlineStr\""))
            {
                var tm = Regex.Match(inner, @"<t>([^<]*)</t>");
                if (tm.Success) text = tm.Groups[1].Value.Trim();
            }
            else
            {
                // Shared string referansı
                var vm = Regex.Match(inner, @"<v>(\d+)</v>");
                if (vm.Success && int.TryParse(vm.Groups[1].Value, out int si) && si < shared.Length)
                    text = shared[si].Trim();
            }

            if (string.IsNullOrEmpty(text)) continue;
            if (!csvByNorm.TryGetValue(ApplyAliases(Normalize(text)), out int csvIdx)) continue;
            if (colRef.Length != 1) continue; // çok harfli sütun (XFD gibi) şimdilik atla

            map.Add((colRef[0], csvIdx));
        }

        return map.ToArray();
    }

    // xl/sharedStrings.xml'den tüm string değerlerini dizi olarak döndürür.
    private static string[] LoadSharedStrings(string zipPath)
    {
        try
        {
            using var zip = new ZipArchive(
                new FileStream(zipPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite),
                ZipArchiveMode.Read);
            var entry = zip.GetEntry("xl/sharedStrings.xml");
            if (entry == null) return Array.Empty<string>();

            string xml;
            using (var sr = new StreamReader(entry.Open(), Encoding.UTF8))
                xml = sr.ReadToEnd();

            var list = new List<string>();
            foreach (Match m in Regex.Matches(xml, @"<si>([\s\S]*?)</si>"))
            {
                // <si> içindeki tüm <t>...</t> parçalarını birleştir (rich text desteği)
                var parts = Regex.Matches(m.Groups[1].Value, @"<t[^>]*>([^<]*)</t>");
                list.Add(string.Concat(parts.Cast<Match>().Select(p => p.Groups[1].Value)));
            }
            return list.ToArray();
        }
        catch { return Array.Empty<string>(); }
    }

    // Türkçe karakter dönüşümü + harf dışı tüm karakterleri sil
    private static string Normalize(string s)
    {
        var up = s.ToUpperInvariant()
                  .Replace("İ", "I").Replace("Ğ", "G").Replace("Ş", "S")
                  .Replace("Ü", "U").Replace("Ö", "O").Replace("Ç", "C");
        return new string(up.Where(c => c >= 'A' && c <= 'Z').ToArray());
    }

    private static readonly Dictionary<string, string> ColAliases =
        new(StringComparer.OrdinalIgnoreCase)
    {
        ["ONCKAPANIS"] = "ONCEKIKAPANIS",
        ["AGIRLORT"]   = "AGIRLIKLIORT",
        ["DEGISIM"]    = "DEGISIMYUZDE",
    };

    private static string ApplyAliases(string norm) =>
        ColAliases.TryGetValue(norm, out string? a) ? a : norm;

    // ── TARİHSEL_VERİ: sadece yeni (ticker, tarih) çiftleri eklenir ─────────────
    public void WriteTarihselVeri(Dictionary<string, List<StockBar>> allData)
    {
        if (!_sheetXmlMap.TryGetValue(SH_TARIHSEL, out var xmlPath)) return;
        string origXml = ReadEntry(xmlPath);

        // Mevcut (ticker|tarih) çiftlerini ve son satır numarasını bul
        var existing = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        int lastRow  = 3;

        foreach (Match m in Regex.Matches(origXml,
            @"<row r=""(\d+)"">[\s\S]*?</row>", RegexOptions.None))
        {
            if (!int.TryParse(m.Groups[1].Value, out int r) || r < 4) continue;
            lastRow = Math.Max(lastRow, r);

            var tM = Regex.Match(m.Value,
                @"<c r=""A\d+"" t=""inlineStr""><is><t>([^<]+)</t></is></c>");
            var dM = Regex.Match(m.Value,
                @"<c r=""B\d+"" t=""inlineStr""><is><t>([^<]+)</t></is></c>");
            if (tM.Success && dM.Success)
                existing.Add($"{tM.Groups[1].Value.Trim()}|{dM.Groups[1].Value.Trim()}");
        }

        // Yeni satırları oluştur
        var sb      = new StringBuilder(2 * 1024 * 1024);
        int row     = lastRow + 1;
        int newRows = 0;

        foreach (var (ticker, bars) in allData)
        {
            var slice = bars.Count > 65 ? bars.Skip(bars.Count - 65).ToList() : bars;
            foreach (var bar in slice)
            {
                string key = $"{bar.Ticker}|{bar.Date:dd.MM.yyyy}";
                if (existing.Contains(key)) continue;

                sb.Append($"<row r=\"{row}\">");
                sb.Append($"<c r=\"A{row}\" t=\"inlineStr\"><is><t>{Esc(bar.Ticker)}</t></is></c>");
                sb.Append($"<c r=\"B{row}\" t=\"inlineStr\"><is><t>{bar.Date:dd.MM.yyyy}</t></is></c>");
                sb.Append($"<c r=\"C{row}\" t=\"n\"><v>{bar.Open.ToString("R", IC)}</v></c>");
                sb.Append($"<c r=\"D{row}\" t=\"n\"><v>{bar.High.ToString("R", IC)}</v></c>");
                sb.Append($"<c r=\"E{row}\" t=\"n\"><v>{bar.Low.ToString("R", IC)}</v></c>");
                sb.Append($"<c r=\"F{row}\" t=\"n\"><v>{bar.Close.ToString("R", IC)}</v></c>");
                sb.Append($"<c r=\"G{row}\" t=\"n\"><v>{bar.Volume}</v></c>");
                sb.Append($"<c r=\"H{row}\" t=\"n\"><v>{bar.WeightedAvg.ToString("R", IC)}</v></c>");
                sb.Append("</row>");
                row++;
                newRows++;
            }
        }

        if (newRows == 0)
        {
            Console.WriteLine("[Excel] TARİHSEL_VERİ: Yeni satır yok, atlandı.");
            return;
        }

        // Yeni satırları </sheetData> öncesine ekle
        int sdClose = origXml.IndexOf("</sheetData>", StringComparison.Ordinal);
        string newXml = sdClose >= 0
            ? origXml[..sdClose] + sb.ToString() + origXml[sdClose..]
            : origXml + sb.ToString();

        // A4:H4 merge varsa kaldır (veri satırlarıyla çakışır)
        newXml = newXml.Replace("<mergeCell ref=\"A4:H4\"/>", "");

        _pending[xmlPath] = Encoding.UTF8.GetBytes(newXml);
        Console.WriteLine($"[Excel] TARİHSEL_VERİ: {newRows} yeni satır eklendi " +
                          $"(toplam ~{row - 4} satır).");
    }


    public bool HasChanges => _pending.Count > 0;

    // ── Kaydet ────────────────────────────────────────────────────────────────
    public void Save()
    {
        if (_pending.Count == 0) return; // hiçbir şey değişmedi, kaydetme

        var result = new MemoryStream();

        using (var src    = new FileStream(_sourcePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
        using (var srcZip = new ZipArchive(src, ZipArchiveMode.Read))
        using (var dstZip = new ZipArchive(result, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach (var entry in srcZip.Entries)
            {
                if (entry.FullName.Equals("xl/calcChain.xml",
                    StringComparison.OrdinalIgnoreCase)) continue;

                var ne = dstZip.CreateEntry(entry.FullName, CompressionLevel.Fastest);
                using var ins  = entry.Open();
                using var outs = ne.Open();

                if (_pending.TryGetValue(entry.FullName, out var newBytes))
                    outs.Write(newBytes, 0, newBytes.Length);
                else
                    ins.CopyTo(outs);
            }
        }

        ClearDataDescriptorBit(result);

        // Önce geçici dosyaya yaz, sonra atomik taşı.
        // Excel çıktıyı açıksa taşıma başarısız olur — kullanıcıyı bilgilendir.
        string tmpPath = _outputPath + ".tmp";
        using (var fs = new FileStream(tmpPath, FileMode.Create, FileAccess.Write))
        {
            result.Position = 0;
            result.CopyTo(fs);
        }

        try
        {
            if (File.Exists(_outputPath)) File.Delete(_outputPath);
            File.Move(tmpPath, _outputPath);
            long sz = new FileInfo(_outputPath).Length;
            Console.WriteLine($"[Excel] Kaydedildi: {_outputPath}  ({sz / 1024} KB)");
        }
        catch (IOException)
        {
            Console.WriteLine($"[Excel] UYARI — Çıktı dosyası Excel'de açık, kapatıp tekrar deneyin.");
            Console.WriteLine($"[Excel] Veri geçici dosyada hazır: {tmpPath}");
        }
    }

    // ── Yardımcı: Kaynak ZIP'ten dosya oku (_sourcePath) ─────────────────────
    private string ReadEntry(string entryPath) => ReadEntryFrom(_sourcePath, entryPath);

    // ── Yardımcı: Her zaman şablondan oku (_templatePath) ────────────────────
    private string ReadEntryFromTemplate(string entryPath) => ReadEntryFrom(_templatePath, entryPath);

    private static string ReadEntryFrom(string zipPath, string entryPath)
    {
        using var zip = new ZipArchive(
            new FileStream(zipPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite), ZipArchiveMode.Read);

        var entry = zip.GetEntry(entryPath)
                 ?? zip.Entries.FirstOrDefault(e =>
                        string.Equals(e.FullName, entryPath, StringComparison.OrdinalIgnoreCase));

        if (entry == null)
            throw new FileNotFoundException($"ZIP içinde bulunamadı: {entryPath}");

        using var sr = new StreamReader(entry.Open(), Encoding.UTF8);
        return sr.ReadToEnd();
    }

    // ── Yardımcı: Sayfa adı → XML yolu haritası ──────────────────────────────
    private static Dictionary<string, string> ReadSheetXmlPaths(string xlsmPath)
    {
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        try
        {
            using var zip = new ZipArchive(
                new FileStream(xlsmPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite), ZipArchiveMode.Read);

            var wbEntry = zip.GetEntry("xl/workbook.xml");
            if (wbEntry == null) return result;
            string wbXml;
            using (var sr = new StreamReader(wbEntry.Open(), Encoding.UTF8))
                wbXml = sr.ReadToEnd();

            var relsEntry = zip.GetEntry("xl/_rels/workbook.xml.rels");
            if (relsEntry == null) return result;
            string relsXml;
            using (var sr = new StreamReader(relsEntry.Open(), Encoding.UTF8))
                relsXml = sr.ReadToEnd();

            var ridToFile = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (Match m in Regex.Matches(relsXml, @"<Relationship\b[^>]*>"))
            {
                var idM     = Regex.Match(m.Value, @"\bId=""([^""]+)""");
                var targetM = Regex.Match(m.Value, @"\bTarget=""([^""]+)""");
                if (idM.Success && targetM.Success)
                {
                    string target = targetM.Groups[1].Value.TrimStart('/');
                    // Target değerleri xl/_rels/ içinden göreceli → xl/ ön eki ekle
                    if (!target.StartsWith("xl/", StringComparison.OrdinalIgnoreCase))
                        target = "xl/" + target;
                    ridToFile[idM.Groups[1].Value] = target;
                }
            }

            foreach (Match m in Regex.Matches(wbXml, @"<sheet\b[^>]*>"))
            {
                var nameM = Regex.Match(m.Value, @"\bname=""([^""]*)""");
                var ridM  = Regex.Match(m.Value, @"\br:id=""([^""]*)""");
                if (nameM.Success && ridM.Success &&
                    ridToFile.TryGetValue(ridM.Groups[1].Value, out var file))
                    result[nameM.Groups[1].Value] = file;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"[Uyarı] Sayfa haritası okunamadı: {ex.Message}");
        }
        return result;
    }

    // ── Yardımcı: .NET ZipArchive bit-3 düzeltmesi ───────────────────────────
    private static void ClearDataDescriptorBit(MemoryStream ms)
    {
        var data = ms.GetBuffer();
        int len  = (int)ms.Length;
        for (int i = 0; i <= len - 4; i++)
        {
            if (data[i] != 0x50 || data[i + 1] != 0x4B) continue;
            if      (data[i+2] == 0x03 && data[i+3] == 0x04 && i+8  < len) data[i+6] &= 0xF7;
            else if (data[i+2] == 0x01 && data[i+3] == 0x02 && i+10 < len) data[i+8] &= 0xF7;
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

    // ── Yardımcı: XML özel karakter kaçışı ───────────────────────────────────
    private static string Esc(string s) =>
        s.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;");

    // ── Yardımcı: Sayısal hücre yaz ──────────────────────────────────────────
    private static void AppendNum(StringBuilder sb, string cellRef, string raw)
    {
        raw = raw.Trim().Replace(',', '.');
        if (double.TryParse(raw, NumberStyles.Any, IC, out double v))
            sb.Append($"<c r=\"{cellRef}\" t=\"n\"><v>{v.ToString("R", IC)}</v></c>");
        else
            sb.Append($"<c r=\"{cellRef}\" t=\"inlineStr\"><is><t>{Esc(raw)}</t></is></c>");
    }
}
