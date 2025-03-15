using System;
using System.Collections.Generic;
using System.Diagnostics;

public class ExcelExportTypeDefine
{
    #region Config
    /// <summary>
    /// 字符串拆分列表
    /// </summary>
    public static List<T> GetList<T>(string str, char spliteChar)
    {
        string[] ss = str.Split(spliteChar);
        int length = ss.Length;
        List<T> arry = new List<T>(ss.Length);
        for (int i = 0; i < length; i++)
        {
            if (!string.IsNullOrEmpty(ss[i]))
            {
                arry.Add(GetValue<T>(ss[i]));
            }
        }
        return arry;
    }

    public static T GetValue<T>(object value)
    {
        return (T)Convert.ChangeType(value, typeof(T));
    }

    public static string GetItemStructStr(string str,bool isArray)
    {
        string strResult = "";
        List<string> list = GetList<string>(str, ',');
        for (int i = 0; i < list.Count; i++)
        {
            List<string> item = GetList<string>(list[i], '_');
            if (item.Count < 2)
            {
                UnityEngine.Debug.LogError("配置错误！" + str);
                return null;
            }
            strResult += "{ID:" + list[i].Replace("_", ",Num:") + "}" + (i + 1 == list.Count ? "" : ",");
        }
        return (isArray ? "[" : "") + strResult + (isArray ? "]" : "");
    }

    public static string GetKeyValStructStr(string str,bool isArray)
    {
        string strResult = "";
        List<string> list = GetList<string>(str, ',');
        for (int i = 0; i < list.Count; i++)
        {
            List<string> item = GetList<string>(list[i], '_');
            if (item.Count < 2)
            {
                UnityEngine.Debug.LogError("配置错误！" + str);
                return null;
            }
            strResult += "{Key:" + list[i].Replace("_", ",Val:") + "}" + (i + 1 == list.Count ? "" : ",");
        }
        return (isArray ? "[" : "") + strResult + (isArray ? "]" : "");
    }
    public static string GetLevelRangeStructStr(string str, bool isArray)
    {
        string strResult = "";
        //60001_1,60002_1
        List<string> list = GetList<string>(str, ',');
        for (int i = 0; i < list.Count; i++)
        {
            List<string> item = GetList<string>(list[i], '_');
            if (item.Count < 2)
            {
                UnityEngine.Debug.LogError("配置错误！" + str);
                return null;
            }
            strResult += "{minLv:" + list[i].Replace("_", ",maxLv:") + "}" + (i + 1 == list.Count ? "" : ",");
        }
        return (isArray ? "[" : "") + strResult + (isArray ? "]" : "");
    }
    #endregion



    #region CS Temp
    public static string GetType(string typeName)
    {
        if (typeName.Contains("ItemStruct")) return typeName;

        else if (typeName.Contains("EDailyTaskType")) return typeName;

        else if (typeName.Contains("KeyVal")) return typeName;
        else if (typeName.Contains("LevelRange")) return typeName;
        return typeName.ToLower();
    }
    #endregion
}

