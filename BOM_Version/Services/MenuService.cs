using SAPbouiCOM;

namespace BOM_Version.Services
{
    public class MenuService
    {
        private readonly string _formUid;

        public MenuService(string formUid)
        {
            _formUid = formUid;
        }

        // ============================================================
        // ADD MENU ITEM
        // ============================================================
        public void AddMenuItem(string parentId, string uid, string title)
        {
            try
            {
                if (!Program.SBO_Application.Menus.Exists(parentId))
                    return;

                var parent = Program.SBO_Application.Menus.Item(parentId);

                if (!parent.SubMenus.Exists(uid))
                {
                    var mcp = (MenuCreationParams)
                        Program.SBO_Application.CreateObject(BoCreatableObjectType.cot_MenuCreationParams);

                    mcp.UniqueID = uid;
                    mcp.Type = BoMenuType.mt_STRING;
                    mcp.String = title;

                    parent.SubMenus.AddEx(mcp);
                }
            }
            catch { /* menu error ignore */ }
        }

        // ============================================================
        // ENABLE / DISABLE NAVIGATION BUTTONS
        // ============================================================
        public void EnableNavigation(bool enabled)
        {
            try
            {
                if (!FormExists()) return;

                var form = Program.SBO_Application.Forms.Item(_formUid);

                form.EnableMenu("1288", enabled);   // First
                form.EnableMenu("1289", enabled);   // Prev
                form.EnableMenu("1290", enabled);   // Next
                form.EnableMenu("1291", enabled);   // Last
            }
            catch
            {
                // avoid showing popup error
            }
        }

        // ============================================================
        // CHECK FORM EXISTS
        // ============================================================
        private bool FormExists()
        {
            try
            {
                var forms = Program.SBO_Application.Forms;
                for (int i = 0; i < forms.Count; i++)
                {
                    if (forms.Item(i).UniqueID == _formUid)
                        return true;
                }
            }
            catch { }
            return false;
        }
    }
}
