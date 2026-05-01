using System;
using System.IO;
using System.Collections.Generic;
using System.Text;

using Matriks.Data.Symbol;
using Matriks.Engines;
using Matriks.Symbols;
using Matriks.Trader.Core;
using Matriks.Trader.Core.Fields;
using Matriks.Lean.Algotrader.AlgoBase;
using Matriks.Lean.Algotrader.Models;

namespace Matriks.Lean.Algotrader
{
	public class BistCanliVeri : MatriksAlgo
	{
		private const string CIKTI_YOLU = @"C:\Users\thbay\OneDrive\Masaüstü\Borsa\BIST_v33_K4K5_FINAL\bist_canli.csv";
		private const int KAYIT_SIKLIGI = 60;

		private readonly string[] HISSELER =
		{
			// PARÇA 1 - 50
			"ACSEL", "ADEL", "ADESE", "ADGYO", "AEFES", "AFYON", "AGROT", "AGYO", "AKENR", "AKFGY", "AKGRT", "AKSA", "AKSEN", "AKSGY", "ALARK", "ALBRK", "ALCAR", "ALCTL", "ALFAS", "ALGYO", "ALKIM", "ANELE", "ANGEN", "ANHYT", "ANSGR", "ARASE", "ARCLK", "ARDYZ", "ARENA", "ARSAN", "ARTMS", "ARZUM", "ASELS", "ASGYO", "ASTOR", "ATAGY", "ATAKP", "ATEKS", "ATLAS", "ATPET", "AVHOL", "AVGYO", "AVOD", "AVTUR", "AYDEM", "AYEN", "AYGAZ", "BAGFS", "BAKAB", "BALAT",
			// PARÇA 2 - 50
"BANVT", "BASGZ", "BAYRK", "BERA", "BFREN", "BIENY", "BIOEN", "BLCYT", "BMSCH", "BNTAS", "BORSK", "BOSSA", "BRISA", "BRKSN", "BRSAN", "BRYAT", "BSOKE", "BTCIM", "BUCIM", "BURCE", "BURVA", "CCOLA", "CELHA", "CEMAS", "CEMTS", "CIMSA", "CLEBI", "CMENT", "CUSAN", "DAGI", "DENGE", "DEVA", "DGATE", "DGNMO", "DITAS", "DMSAS", "DOAS", "DOHOL", "DOKTA", "DYOBY", "DZGYO", "ECILC", "ECZYT", "EDATA", "EDIP", "EGEEN", "EGGUB", "EGPRO", "EGSER", "EKGYO",
			// PARÇA 3 - 50
"ELITE", "EMKEL", "EMNIS", "ENERY", "ENKAI", "EPLAS", "ERBOS", "EREGL", "ERSU", "ESEN", "ESEMS", "ETILR", "EUYO", "EYGYO", "FENER", "FLAP", "FONET", "FORTE", "FROTO", "GARAN", "GEDIK", "GEDZA", "GENIL", "GEREL", "GLBMD", "GLYHO", "GMTAS", "GOLTS", "GOODY", "GRSEL", "GSDDE", "GSDHO", "GSRAY", "GUBRF", "HALKB", "HATEK", "HDFGS", "HEDEF", "HEKTS", "HKTM", "HTTBT", "HUBVC", "HUNER", "HURGZ", "ICBCT", "IDGYO", "IHEVA", "IHGZT", "IHLGM", "IHLAS",
			// PARÇA 4 - 50
"IMASM", "INDES", "INFO", "INTEM", "INVEO", "ISATR", "ISBIR", "ISCTR", "ISFIN", "ISGSY", "ISGYO", "ISKUR", "IZENR", "IZFAS", "IZINV", "IZMDC", "JANTS", "KAPLM", "KAREL", "KARSN", "KARTN", "KLGYO", "KLMSN", "KLSER", "KNFRT", "KONYA", "KORDS", "KRDMA", "KRDMB", "KRDMD", "KRONT", "KRSTL", "KRTEK", "KRVGD", "KTLEV", "KUTPO", "LIDER", "LILAK", "LINK", "LUKSK", "MARKA", "MARTI", "MAVI", "MEDTR", "MERIT", "MERKO", "METRO", "MGROS", "MNDRS", "MPARK",
			// PARÇA 5 - 50
"MRSHL", "MSGYO", "NATEN", "NETAS", "NIBAS", "NTHOL", "NUHCM", "OBAMS", "ODINE", "ONCSM", "ORCAY", "ORGE", "ORMA", "OSMEN", "OTKAR", "OTTO", "OYAKC", "OYYAT", "OZGYO", "PAGYO", "PAMEL", "PENGD", "PENTA", "PETKM", "PGSUS", "PINSU", "POLTK", "PRZMA", "RALYH", "RAYSG", "RGYAS", "RYGYO", "SAFKR", "SAMAT", "SANEL", "SANFM", "SANKO", "SARKY", "SAYAS", "SEKFK", "SEKUR", "SELEC", "SILVR", "SKBNK", "SKTAS", "SMART", "SNGYO", "SOKM", "SONME", "SUMAS",
			// PARÇA 6 - 48
"SUNTK", "SUWEN", "TABGD", "TATGD", "TAVHL", "TBORG", "TCELL", "TDGYO", "THYAO", "TKNSA", "TLMAN", "TOASO", "TRILC", "TSGYO", "TSKB", "TSPOR", "TTKOM", "TTRAK", "TUCLK", "TURSG", "TUPRS", "ULUFA", "ULUSE", "UNLU", "USAK", "VAKBN", "VAKFN", "VAKKO", "VERUS", "VESBE", "VESTL", "YATAS", "YAYLA", "YKBNK", "YONGA", "YUNSA", "ZEDUR", "ZRGYO", "ZOREN", "AKBNK", "AKCNS", "AGESA", "BIMAS", "SASA", "KCHOL", "SAHOL", "TKFEN", "AGHOL",

		};

		[SymbolParameter("GARAN")]
		public string AnaSembol;

		[Parameter(SymbolPeriod.Min)]
		public SymbolPeriod Periyot;

		public override void OnInit()
		{
			Debug("Canlı Piyasa Kaydedici Başlatılıyor...");

			foreach (var hisse in HISSELER)
			{
				AddSymbol(hisse, Periyot);
			}

			SetTimerInterval(KAYIT_SIKLIGI);
			WorkWithPermanentSignal(true);

			Debug($"Sistem Hazır! {HISSELER.Length} adet hisse her {KAYIT_SIKLIGI} saniyede bir CSV'ye kaydedilecek.");
		}

		public override void OnDataUpdate(BarDataEventArgs barData)
		{
		}

		public override void OnTimer()
		{
			List<string> anlikSatirlar = new List<string>();

			int no = 1;

			foreach (var hisse in HISSELER)
			{
				try
				{
					var bd = GetBarData(hisse, Periyot);

					if (bd == null || bd.Close == null || bd.Close.Count == 0)
						continue ;

					int son = bd.Close.Count - 1;

					decimal sonFiyat = bd.Close[son];
					decimal acilis = bd.Open[son];
					decimal yuksek = bd.High[son];
					decimal dusuk = bd.Low[son];
					decimal hacimLot = bd.Volume[son];

					if (sonFiyat <= 0)
						continue ;

					decimal oncekiKapanis = son > 0 ? bd.Close[son - 1] : sonFiyat;

					decimal degisimYuzde = oncekiKapanis > 0
						? ((sonFiyat - oncekiKapanis) / oncekiKapanis) * 100
						: 0;

					decimal agirlikliOrt = (acilis + yuksek + dusuk + sonFiyat) / 4;
					decimal degerTL = sonFiyat * hacimLot;

					string temizHisse = hisse.Replace(".E", "").Trim();

					string satir =
						no + ";" +
						temizHisse + ";" +
						FormatDecimal(sonFiyat) + ";" +
						FormatDecimal(acilis) + ";" +
						FormatDecimal(yuksek) + ";" +
						FormatDecimal(dusuk) + ";" +
						FormatDecimal(oncekiKapanis) + ";" +
						FormatDecimal(hacimLot) + ";" +
						FormatDecimal(degisimYuzde) + ";" +
						FormatDecimal(agirlikliOrt) + ";" +
						FormatDecimal(degerTL);

					anlikSatirlar.Add(satir);
					no++;
				}
				catch
				{
				}
			}

			if (anlikSatirlar.Count > 0)
			{
				YazCSV(anlikSatirlar);
			}
		}

		private void YazCSV(List<string> satirlar)
		{
			try
			{
				string klasor = Path.GetDirectoryName(CIKTI_YOLU);

				if (!Directory.Exists(klasor))
					Directory.CreateDirectory(klasor);

				using (var sw = new StreamWriter(CIKTI_YOLU, false, Encoding.UTF8))
				{
					sw.WriteLine("NO;HISSE;SON_FIYAT;ACILIS;YUKSEK;DUSUK;ONCEKI_KAPANIS;HACIM_LOT;DEGISIM_YUZDE;AGIRLIKLI_ORT;DEGER_TL");

					foreach (var s in satirlar)
					{
						sw.WriteLine(s);
					}
				}

				Debug($"[ {DateTime.Now:HH:mm:ss} ] - {satirlar.Count} hissenin canlı verisi CSV'ye yazıldı.");
			}
			catch (Exception ex)
			{
				Debug("HATA Canlı Veri Yazılamadı: " + ex.Message);
			}
		}

		private string FormatDecimal(decimal value)
		{
			return value.ToString("0.####").Replace(",", ".");
		}

		public override void OnStopped()
		{
			Debug("Canlı Kayıt Sistemi Durduruldu.");
		}
	}
}
