using ClosedXML.Excel;
using Microsoft.Win32;
using System.Linq;
using System.Windows;

namespace Excel2DataModel
{
	/// <summary>
	/// Dialog.xaml の相互作用ロジック
	/// </summary>
	public partial class Dialog : Window
	{
		public Dialog(EA.Repository repository)
		{
			InitializeComponent();

			// 参照ボタン
			BtnReference.Click += (s, e) =>
			{
				var dialog = new OpenFileDialog();
				dialog.Filter = "エクセルファイル（*.xlsx）|*.xlsx";

				if (dialog.ShowDialog() == true)
				{
					TbxFileName.Text = dialog.FileName;
				}
			};

			// OKボタン
			BtnOk.Click += (s, e) =>
			{
				if (string.IsNullOrWhiteSpace(TbxFileName.Text)) return;
				var book = new XLWorkbook(TbxFileName.Text);

				// パッケージ作成
				var model = (EA.Package)repository.Models.GetAt(0);
				var package = (EA.Package)model.Packages.AddNew($"{TbxFunctionName.Text}(論理データモデル図)", "");
				package.Update();

				// ダイアグラム作成
				var diagram = (EA.Diagram)package.Diagrams.AddNew(TbxFunctionName.Text, "Logical");
				diagram.MetaType = "Extended::Data Modeling";
				diagram.Version = "1.0";
				diagram.Update();

				// テーブル作成
				foreach (var sheet in book.Worksheets)
				{
					if (sheet.Position == 1) continue;
					if (sheet.Visibility != XLWorksheetVisibility.Visible) continue;

					// テーブルエレメント作成
					var element = (EA.Element)package.Elements.AddNew(sheet.Cell("C6").Value.ToString(), "Table");
					element.Stereotype = "table";
					element.Gentype = "Oracle";
					element.Update();

					// カラム定義の最初の行を取得
					var baseRow = sheet.Search("カラム定義").FirstOrDefault();
					if (baseRow == null) continue;
					var currentRow = baseRow.Address.RowNumber + 2;

					// テーブル定義の行を論理データモデル図に出力
					var position = 0;
					while (true)
					{
						// 論理名のセルを見て出力するか否かを判定
						if (string.IsNullOrWhiteSpace(sheet.Cell(currentRow, 2).Value.ToString())) break;

						// 属性作成
						var attribute = (EA.Attribute)element.Attributes.AddNew(sheet.Cell(currentRow, 2).Value.ToString(), sheet.Cell(currentRow, 4).Value.ToString());
						attribute.Stereotype = "column";

						// PK
						if (sheet.Cell(currentRow, 3).Value.ToString() == "ID") attribute.IsOrdered = true;

						// 位置
						attribute.Pos = position;

						// 別名
						attribute.Alias = sheet.Cell(currentRow, 3).Value.ToString();

						if (attribute.Type == "VARCHAR2" || attribute.Type == "CHAR") attribute.Length = sheet.Cell(currentRow, 5).Value.ToString();

						// サイズと小数桁数
						if (attribute.Type == "NUMBER")
						{
							attribute.Precision = sheet.Cell(currentRow, 5).Value.ToString();
							attribute.Scale = "0";
						}

						// 初期値
						attribute.Default = sheet.Cell(currentRow, 7).Value.ToString();

						// 必須
						if (sheet.Cell(currentRow, 10).Value.ToString() == "○" ||
							sheet.Cell(currentRow, 10).Value.ToString() == "〇") attribute.AllowDuplicates = true;

						// 更新
						attribute.Update();

						currentRow++;
						position++;
					}

					element.Attributes.Refresh();

					var key = (EA.Attribute)element.Attributes.GetAt(0);
					if (key == null) continue;
					if (key.Alias == "ID")
					{
						// 制約作成
						var constraint = (EA.Constraint)element.Constraints.AddNew($"PK_{sheet.Cell("C6").Value}", "PK");
						constraint.Weight = key.AttributeID;
						constraint.Update();
					}
				}

				Close();
			};
		}
	}
}
