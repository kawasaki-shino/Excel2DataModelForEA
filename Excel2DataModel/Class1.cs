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
	}
}
