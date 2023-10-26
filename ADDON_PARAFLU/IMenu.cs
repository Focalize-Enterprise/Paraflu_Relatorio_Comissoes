using SAPbouiCOM;

namespace ADDON_PARAFLU
{
    public interface IMenu
    {
        void AddMenuItems();
        void RemoveMenus();
        void SBO_Application_MenuEvent(ref MenuEvent pVal, out bool BubbleEvent);
    }
}