using System;
using System.IO;
using System.Collections.Generic;
using System.Text;
using System.Linq;

using Matriks.Data.Symbol;
using Matriks.Engines;
using Matriks.Symbols;
using Matriks.Trader.Core;
using Matriks.Trader.Core.Fields;
using Matriks.Lean.Algotrader.AlgoBase;
using Matriks.Lean.Algotrader.Models;

namespace Matriks.Lean.Algotrader
{
	public class BistTarihselVeri : MatriksAlgo
	{
		private const string CIKTI_YOLU =
			@"C:\Users\thbay\OneDrive\Masaüstü\Borsa\BIST_v33_K4K5_FINAL\bist_tarihsel.csv";

		private const int BAR_SAYISI = 75;

		[SymbolParameter("GARAN")]
		public string Symbol;

		[Parameter(SymbolPeriod.Day)]
		public SymbolPeriod SymbolPeriod;

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
"SUNTK", "SUWEN", "TABGD", "TATGD", "TAVHL", "TBORG", "TCELL", "TDGYO", "THYAO", "TKNSA", "TLMAN", "TOASO", "TRILC", "TSGYO", "TSKB", "TSPOR", "TTKOM", "TTRAK", "TUCLK", "TURSG", "TUPRS", "ULUFA", "ULUSE", "UNLU", "USAK", "VAKBN", "VAKFN", "VAKKO", "VERUS", "VESBE", "VESTL", "YATAS", "YAYLA", "YKBNK", "YONGA", "YUNSA", "ZEDUR", "ZRGYO", "ZOREN", "AKBNK", "AKCNS", "AGESA", "BIMAS", "SASA", "KCHOL", "SAHOL", "TKFEN", "AGHOL"
		};

		private readonly List<string> _satirlar = new List<string>();
		private bool _yazildi = false;

		public override void OnInit()
		{
			Debug("OnInit çalıştı");
			Debug("Ana Symbol: " + Symbol);
			Debug("SymbolPeriod: " + SymbolPeriod);
			Debug("Toplam hisse: " + HISSELER.Distinct().Count());

			foreach (var hisse in HISSELER.Distinct())
			{
				AddSymbol(hisse, SymbolPeriod);
			}

			SetTimerInterval(1);
			WorkWithPermanentSignal(true);

			Debug("Tüm hisseler AddSymbol tamamlandı");
			Debug("Timer başlatıldı. Veri hazır olunca CSV yazılacak.");
		}

		public override void OnDataUpdate(BarDataEventArgs barData)
		{
			// Boş bırakıyoruz. Yazmayı OnTimer yapacak.
		}

		public override void OnTimer()
		{
			if (_yazildi)
				return;

			Debug("Veri kontrol ediliyor...");

			try
			{
				_satirlar.Clear();

				foreach (var hisse in HISSELER.Distinct())
				{
					try
					{
						var bd = GetBarData(hisse, SymbolPeriod);

						if (bd == null || bd.BarDataIndexer == null)
						{
							Debug("Bar data boş: " + hisse);
							continue ;
						}

						int sonIdx = bd.BarDataIndexer.LastBarIndex;

						if (sonIdx <= 0)
						{
							Debug("Yeterli bar yok: " + hisse);
							continue ;
						}

						int baslangic = Math.Max(0, sonIdx - BAR_SAYISI + 1);
						string sembol = hisse.Replace(".E", "").Trim();

						int hisseSatirSayisi = 0;

						for (int i = baslangic; i <= sonIdx; i++)
						{
							DateTime tarih = bd.Time[i];
							decimal acilis = bd.Open[i];
							decimal yuksek = bd.High[i];
							decimal dusuk = bd.Low[i];
							decimal kapanis = bd.Close[i];
							decimal hacim = bd.Volume[i];

							if (kapanis <= 0)
								continue ;

							decimal agOrt = (yuksek + dusuk + kapanis) / 3;

							string satir = string.Format(
								"{0};{1};{2:F4};{3:F4};{4:F4};{5:F4};{6};{7:F4}",
								sembol,
								tarih.ToString("dd.MM.yyyy"),
								acilis,
								yuksek,
								dusuk,
								kapanis,
								hacim,
								agOrt
							);

							_satirlar.Add(satir);
							hisseSatirSayisi++;
						}

						Debug("Tamamlandı: " + sembol + " | Satır: " + hisseSatirSayisi);
					}
					catch (Exception ex)
					{
						Debug("Hisse hata: " + hisse + " -> " + ex.Message);
					}
				}

				if (_satirlar.Count == 0)
				{
					Debug("Henüz veri oluşmadı. Timer tekrar deneyecek.");
					return;
				}

				_yazildi = true;
				Debug("Toplam satır: " + _satirlar.Count);

				YazCSV();
			}
			catch (Exception ex)
			{
				Debug("OnTimer hata: " + ex.Message);
			}
		}

		private void YazCSV()
		{
			try
			{
				string klasor = Path.GetDirectoryName(CIKTI_YOLU);

				if (!Directory.Exists(klasor))
					Directory.CreateDirectory(klasor);

				bool dosyaVarMi = File.Exists(CIKTI_YOLU);

				using (var sw = new StreamWriter(CIKTI_YOLU, true, Encoding.UTF8))
				{
					if (!dosyaVarMi)
						sw.WriteLine("HISSE;TARIH;ACILIS;YUKSEK;DUSUK;KAPANIS;HACIM;A_ORT");

					foreach (var satir in _satirlar)
						sw.WriteLine(satir);
				}

				Debug("CSV kaydedildi: " + CIKTI_YOLU);
			}
			catch (Exception ex)
			{
				Debug("CSV yazma hata: " + ex.Message);
			}
		}

		public override void OnStopped()
		{
			Debug("Sistem durduruldu.");
		}
	}
}
