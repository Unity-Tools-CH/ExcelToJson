using Newtonsoft.Json.Serialization;
using UnityEditor;

public class CfgMgr
{
    public static GameItem GameItem = null;
    public static EquipmentCreate EquipmentCreate = null;

    public static void Init()
    {
        GameItem = new();
        GameItem.LoadData();

        EquipmentCreate = new();
        EquipmentCreate.LoadData();
    }
}
