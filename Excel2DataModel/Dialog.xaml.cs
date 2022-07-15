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
				diagram.StyleEx =
					"ExcludeRTF=0;DocAll=0;HideQuals=0;AttPkg=1;ShowTests=1;ShowMaint=1;SuppressFOC=1;MatrixActive=0;SwimlanesActive=1;KanbanActive=0;MatrixLineWidth=1;MatrixLineClr=0;MatrixLocked=0;TConnectorNotation=Information Engineering;TExplicitNavigability=0;AdvancedElementProps=1;AdvancedFeatureProps=1;AdvancedConnectorProps=1;m_bElementClassifier=1;SPT=1;MDGDgm=Extended::Data Modeling;STBLDgm=;ShowNotes=1;VisibleAttributeDetail=0;ShowOpRetType=1;SuppressBrackets=0;SuppConnectorLabels=0;PrintPageHeadFoot=0;ShowAsList=0;SuppressedCompartments=;Theme=:119;SD=1;SR=1;SRES=1;SIA=1;SIO=1;SIT=1;SFQT=1;SIR=1;SIC=1;ShowProject=1;ShowScenarios=1;SaveTag=E8B0C7D1;";
				diagram.Update();

				// テーブル作成
				foreach (var sheet in book.Worksheets)
				{
					if (sheet.Position == 1) continue;
					if (sheet.Visibility != XLWorksheetVisibility.Visible) continue;

					// テーブルエレメント作成
					var element = (EA.Element)package.Elements.AddNew(sheet.Cell("C6").Value.ToString().Trim().Trim('\u200B'), "Table");
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
						if (string.IsNullOrWhiteSpace(sheet.Cell(currentRow, 2).Value.ToString().Trim().Trim('\u200B'))) break;

						// 属性作成
						var attribute = (EA.Attribute)element.Attributes.AddNew(sheet.Cell(currentRow, 2).Value.ToString().Trim().Trim('\u200B'), sheet.Cell(currentRow, 4).Value.ToString().Trim().Trim('\u200B'));
						attribute.Stereotype = "column";

						// PK
						if (sheet.Cell(currentRow, 3).Value.ToString().Trim().Trim('\u200B') == "ID") attribute.IsOrdered = true;

						// FK
						if (sheet.Cell(currentRow, 3).Value.ToString().Trim().Trim('\u200B') == "REC_ID") attribute.IsCollection = true;

						// 位置
						attribute.Pos = position;

						// 別名
						attribute.Alias = sheet.Cell(currentRow, 3).Value.ToString().Trim().Trim('\u200B');

						if (attribute.Type == "VARCHAR2" || attribute.Type == "CHAR") attribute.Length = sheet.Cell(currentRow, 5).Value.ToString().Trim().Trim('\u200B');

						// サイズと小数桁数
						if (attribute.Type == "NUMBER")
						{
							attribute.Precision = sheet.Cell(currentRow, 5).Value.ToString().Trim().Trim('\u200B');
							attribute.Scale = "0";
						}

						// 初期値
						attribute.Default = sheet.Cell(currentRow, 7).Value.ToString().Trim().Trim('\u200B');

						// 必須
						if (sheet.Cell(currentRow, 10).Value.ToString().Trim().Trim('\u200B') == "○" ||
							sheet.Cell(currentRow, 10).Value.ToString().Trim().Trim('\u200B') == "〇") attribute.AllowDuplicates = true;

						// 更新
						attribute.Update();

						currentRow++;
						position++;
					}

					element.Attributes.Refresh();
					package.Elements.Refresh();

					// PK制約作成
					var key = element.Attributes.Cast<EA.Attribute>().FirstOrDefault(n => n.Alias == "ID");
					if (key != null)
					{
						// 制約の名前と種類を作成
						var method = (EA.Method)element.Methods.AddNew($"PK_{sheet.Cell("C6").Value.ToString().Trim().Trim('\u200B')}", "");
						method.Stereotype = "PK";
						method.Concurrency = "Sequential";
						method.Update();

						// 制約対象を作成
						var param = (EA.Parameter)method.Parameters.AddNew(key.Name, key.Type);
						param.Alias = key.Alias;
						param.Kind = "in";
						param.Update();

						method.Parameters.Refresh();
						method.Update();
					}

					element.Methods.Refresh();

					// ジャーナルのFK制約作成
					if (element.Name.Contains("ジャーナル"))
					{
						// 前提条件をチェック
						// FKの参照先を探す
						var parentElement =
							package.Elements.Cast<EA.Element>().FirstOrDefault(n => n.Name == element.Name.Substring(0, element.Name.Length - 6));
						if (parentElement == null) continue;

						var parentMethod = parentElement.Methods.Cast<EA.Method>()
							.FirstOrDefault(n => n.Stereotype == "PK");
						if (parentMethod == null) continue;

						// FKとなる列
						var fkAttribute = element.Attributes.Cast<EA.Attribute>()
							.FirstOrDefault(n => n.Alias == "REC_ID");
						if (fkAttribute == null) continue;

						// 制約の名前と種類を作成
						var method = (EA.Method)element.Methods.AddNew(
							$"FK_{element.Name}_{parentElement.Name}", "");
						method.Stereotype = "FK";
						method.Concurrency = "Sequential";
						method.StyleEx = $"FKIDX={parentMethod.MethodID}";
						method.Update();

						var param = (EA.Parameter)method.Parameters.AddNew(fkAttribute.Name, fkAttribute.Type);
						param.Alias = fkAttribute.Alias;
						param.Kind = "in";
						param.Update();

						method.Parameters.Refresh();
						method.Update();
						element.Methods.Refresh();

						// コネクターの追加
						var connector = (EA.Connector)element.Connectors.AddNew("AAA", "Association");
						connector.Stereotype = "FK";
						connector.Direction = "Source -> Destination";
						connector.StyleEx = $"FKINFO=SRC={element.Name}:DST={parentElement.Name}";


						// ソース側設定
						connector.ClientID = element.ElementID;
						connector.ClientEnd.Cardinality = "1..*";
						connector.ClientEnd.Constraint = "Unspecified";
						connector.ClientEnd.IsChangeable = "none";
						connector.ClientEnd.Navigable = "Unspecified";
						connector.ClientEnd.Role = element.Methods.Cast<EA.Method>()
							.FirstOrDefault(n => n.Stereotype == "FK")?.Name;

						// ターゲット側設定
						connector.SupplierID = parentElement.ElementID;
						connector.SupplierEnd.Cardinality = "1";
						connector.SupplierEnd.Constraint = "Unspecified";
						connector.SupplierEnd.IsChangeable = "none";
						connector.SupplierEnd.Navigable = "Navigable";
						connector.SupplierEnd.Role = parentElement.Methods.Cast<EA.Method>()
							.FirstOrDefault(n => n.Stereotype == "PK")?.Name;

						connector.Update();
						element.Connectors.Refresh();
					}
				}
				Close();
			};
		}
	}
}
