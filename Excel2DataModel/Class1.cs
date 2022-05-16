namespace Excel2DataModel
{
	public class Class1
	{
		public object EA_GetMenuItems(EA.Repository repository, string menuLocation, string menuName)
		{
			return "論理データモデル図出力";
		}

		public void EA_MenuClick(EA.Repository repository, string menuLocation, string menuName, string itemName)
		{
			var dialog = new Dialog(repository).ShowDialog();
		}

		/// <summary>
		/// 選択中の要素を取得（デバッグ用）
		/// </summary>
		/// <param name="repository"></param>
		/// <param name="menuLocation"></param>
		/// <param name="menuName"></param>
		/// <param name="itemName"></param>
		/// <returns></returns>
		public EA.Element GetSeletedElement(EA.Repository repository, string menuLocation, string menuName, string itemName)
		{
			var diagram = repository.GetCurrentDiagram();
			var diagramObject = (EA.DiagramObject)diagram.SelectedObjects.GetAt(0);
			var ret = repository.GetElementByID(diagramObject.ElementID);

			return ret;
		}
	}
}
