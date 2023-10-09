using Newtonsoft.Json;
using System;
using System.IO;
using System.Windows.Forms;

namespace PowerPointEx {
	/// <summary>設定情報</summary>
	/// <remarks>座標系の単位はポイント</remarks>
	class Config {
		/// <summary>読み込み元のPowerPointのファイルパス</summary>
		public string Input = null;
		/// <summary>目次のスライド名</summary>
		public string TocName = null;
		/// <summary>目次の開始位置</summary>
		public float X = 0;
		/// <summary>目次の開始位置</summary>
		public float Y = 0;
		/// <summary>目次の横幅</summary>
		public float W = 0;
		/// <summary>目次の縦幅</summary>
		public float H = 0;
		/// <summary>保存先のPowerPointのファイルパス(省略すると読み込みファイルに上書きする)</summary>
		public string Output= null;

		/// <summary>保存ファイルパス</summary>
		[JsonIgnore]
		public string SaveName => string.IsNullOrEmpty(Output) ? Input : Output;

		/// <summary>デフォルト値</summary>
		public static Config Default = new Config() {
			Input	= "document.pptx",
			TocName	= "目次",
			Output	= null,
			X		= 10,
			Y		= 30,
			W		= 500,
			H		= 10
		};
	}
	/// <summary>エントリーポイントクラス</summary>
	static class Program {
		/// <summary>設定ファイル名</summary>
		const string ConfigName = "PowerPointToc.json";
		/// <summary>アプリケーションのメイン エントリ ポイントです</summary>
		[STAThread]
		static void Main(string[] Args) {
			try {

				// 設定ファイル
				var Cfg = GetConfig(Args);
				using(var ppt = new PowerPointToc(Cfg.Input)) {
					ppt.UpdateToc(Cfg.TocName, Cfg.X, Cfg.Y, Cfg.W, Cfg.H);
					ppt.Save(Cfg.SaveName);
				}

			} catch(Exception E) {
				MessageBox.Show(E.Message, "エラー");
			}

		}
		/// <summary>設定を取得する</summary>
		/// <param name="Args">引数</param>
		/// <returns>設定</returns>
		private static Config GetConfig(string[] Args) {
			Config Cfg;
			// コマンドラインオプションから設定を作成
			if (Args.Length > 0) {
				Cfg = Config.Default;
				if (Args.Length > 0) Cfg.Input	 = Args[0];
				if (Args.Length > 1) Cfg.TocName = Args[1];
				if (Args.Length > 2) Cfg.X		 = float.Parse(Args[2]);
				if (Args.Length > 3) Cfg.Y		 = float.Parse(Args[3]);
				if (Args.Length > 4) Cfg.W		 = float.Parse(Args[4]);
				if (Args.Length > 5) Cfg.H		 = float.Parse(Args[5]);
				if (Args.Length > 6) Cfg.Output	 = Args[6];
				return Cfg;
			}
			// 設定ファイルの読み込み
			if (File.Exists(ConfigName)) {
				var Text = File.ReadAllText(ConfigName);
				return JsonConvert.DeserializeObject<Config>(Text);
			}
			return Config.Default;
		}
	}
}
