using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.IO;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace PowerPointEx {
	/// <summary>PowerPointTocクラスは、PowerPointファイルの目次を作成または更新します</summary>
	public class PowerPointToc : IDisposable {
		/// <summary>目次のシェイプ目</summary>
		const string TocShapeName = "_toc_shape_";
		/// <summary>パワポのアプリインスタンス</summary>
		private Application app;
		/// <summary>パワポファイルのインスタンス</summary>
		private Presentation doc;
		/// <summary>コンストラクタ</summary>
		/// <param name="filePath">パワポのファイルパス</param>
		public PowerPointToc(string filePath) {
			app = new Application();
			doc = app.Presentations.Open(Path.GetFullPath(filePath));
		}
		/// <summary>目次スライドを取得する</summary>
		/// <param name="Title">スライドタイトル</param>
		/// <returns></returns>
		private Slide GetTocSlide(string Title) {
			// 目次スライドを探す
			foreach (Slide slide in doc.Slides) {
				if (slide.Shapes.HasTitle != MsoTriState.msoTrue) continue;
				if (slide.Shapes.Title.TextFrame.TextRange.Text == Title) return slide;
			}
			return null;
		}
		/// <summary>目次を更新</summary>
		/// <param name="TocSlideName">目次のスライド名</param>
		/// <param name="Xpt">目次の開始位置</param>
		/// <param name="Ypt">目次の開始位置</param>
		/// <param name="Wpt">目次の横幅</param>
		/// <param name="Hpt">目次の縦幅</param>
		public void UpdateToc(string TocSlideName, float Xpt, float Ypt, float Wpt, float Hpt) {
			UpdateToc(TocSlideName, Xpt, Ypt, Wpt, Hpt, DefaultStyle);
		}
		/// <summary>目次を更新</summary>
		/// <param name="TocSlideName">目次のスライド名</param>
		/// <param name="Xpt">目次の開始位置</param>
		/// <param name="Ypt">目次の開始位置</param>
		/// <param name="Wpt">目次の横幅</param>
		/// <param name="Hpt">目次の縦幅</param>
		/// <param name="CustomStyle">テーブルの書式設定用コールバック</param>
		public void UpdateToc(string TocSlideName, float Xpt, float Ypt, float Wpt, float Hpt, Action<Shape> CustomStyle = null) {

			// 目次スライドを探す
			var toc = GetTocSlide("目次");
			if (toc == null) throw new Exception("目次スライドが見つかりませんでした。");

			// 既存の目次を削除
			RemoveTocShape(toc);

			// 目次の表を作成
			var slides = new List<Slide>();
			for (int no = toc.SlideIndex + 1; no < doc.Slides.Count + 1; no++) {
				var slide = doc.Slides[no];
				if (slide.Shapes.HasTitle != MsoTriState.msoTrue) continue;
				slides.Add(slide);
			}
			var shape = toc.Shapes.AddTable(slides.Count, 2, Xpt, Ypt, Wpt, Hpt);
			shape.Name = TocShapeName;

			// 目次の内容を作成
			var table = shape.Table;
			var row = 1;
			foreach (var Slide in slides) {
				table.Cell(row, 1).Shape.TextFrame.TextRange.Text = Slide.Shapes.Title.TextFrame.TextRange.Text;
				table.Cell(row, 2).Shape.TextFrame.TextRange.Text = Slide.SlideNumber.ToString();
				row++;
			}
			// 目次の書式を設定
			if (CustomStyle != null) CustomStyle(shape);
		}
		/// <summary>目次のシェイプを削除する</summary>
		/// <param name="toc"></param>
		private void RemoveTocShape(Slide toc) {
			foreach (Shape shape in toc.Shapes) {
				if (shape.Name != TocShapeName) continue;
				shape.Delete();
			}
		}
		/// <summary>目次の書式を設定</summary>
		/// <param name="shape">目次のシェイプ</param>
		private void DefaultStyle(Shape shape) {
			if (shape == null || shape.HasTable != MsoTriState.msoTrue) return;
			// 行処理
			var table = shape.Table;
			foreach (Row row in table.Rows) {
				// 列処理
				foreach (Cell cell in row.Cells) {
					
					cell.Shape.TextFrame.TextRange.Font.Size = 12;						// フォントサイズを変更
					cell.Shape.TextFrame.TextRange.Font.Bold = MsoTriState.msoFalse;	// フォントサイズを変更
					cell.Shape.TextFrame.TextRange.Font.Color.RGB = 0x000000;			// フォントカラーを変更

					cell.Shape.Fill.Transparency = 1.0f;								// 背景を透明に設定（完全に透明）

					// セル下の罫線
					var border = cell.Borders[PpBorderType.ppBorderBottom];
					border.Visible			= MsoTriState.msoTrue;
					border.Weight			= 1;
					border.ForeColor.RGB	= 0x000000;
				}
				// 右寄せ
				row.Cells[2].Shape.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignRight;

			}
		}
		/// <summary>保存する</summary>
		/// <param name="FileName">ファイル名</param>
		public void Save(string FileName = null) {
			if (FileName != null || !string.IsNullOrEmpty(FileName)) {
				doc.SaveAs(Path.GetFullPath(FileName));
			} else {
				doc.Save();
			}
		}
		/// <summary>変更を保存し、リソースを解放します</summary>
		public void Dispose() {
			doc?.Close();
			app?.Quit();
			doc = null;
			app = null;
		}
	}
}
