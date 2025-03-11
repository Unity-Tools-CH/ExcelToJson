using System.Collections.Generic;
using UnityEngine;
using System.Data;
using System.IO;
using Excel;
using Newtonsoft.Json.Linq;
using System;
using UnityEditor;
using Unity.VisualScripting;
using System.Xml.Linq;

public class ExcelExport
{
    //表格的存放位置
    public static string ConfigPath = Application.dataPath.Replace("Assets", "Excel");

    //public static string DataPath = Application.dataPath.Replace("Assets", "DataConfig/Data/");
    public static string DataPath = Path.Combine(Application.dataPath, "DataConfig/Class/");
    public static string BasePath = Path.Combine(Application.dataPath, "DataConfig/Base/");

    // 实体类模板存放位置
    static string scriptsPath = DataPath;

    // json文件存放位置
    static string jsonPath = Path.Combine(Application.dataPath, "Resources/GameRes/CFG/");



    // 所有表格数据
    static List<TableData> dataList = new List<TableData>();

    [UnityEditor.MenuItem("[GameConfig]/ExportToJson", false, 1)]
    public static void ReadExcel()
    {
        if (!Directory.Exists(DataPath))
        {
            Directory.CreateDirectory(DataPath);
        }

        // 确保 CfgMgr 和 CfgBase 存在
        CreateCfgBase();
        CreateCfgMgr();

        if (Directory.Exists(ConfigPath))
        {
            DirectoryInfo direction = new DirectoryInfo(ConfigPath);
            FileInfo[] files = direction.GetFiles("*", SearchOption.AllDirectories);
            Debug.Log("fileCount:" + files.Length);

            List<string> classList = new List<string>(); // 记录所有导出的类名

            for (int i = 0; i < files.Length; i++)
            {
                if (files[i].Name.StartsWith("~")) continue;
                if (files[i].Name.EndsWith(".meta") || !files[i].Name.EndsWith(".xlsx")) continue;

                string className = files[i].Name.Replace(".xlsx", ""); // 获取类名
                classList.Add(className);

                Debug.Log($"<color=#00ff00>Exporting:</color> {files[i].FullName}");
                LoadData(files[i].FullName, files[i].Name);
                Debug.Log($"<color=#00ff00>Export Finished:</color> {files[i].FullName}");

                AssetDatabase.Refresh();
            }

            // 更新 CfgMgr，追加新类
            UpdateCfgMgr(classList);

            Debug.Log("<color=#00ff00>Export Config Finish!</color>");
        }
        else
        {
            Debug.LogError("ReadExcel configPath not Exists!");
        }
    }

    static void CreateCfgBase()
    {
        string dirPath = Path.Combine(DataPath, "Base");
        if (!Directory.Exists(dirPath))
        {
            Directory.CreateDirectory(dirPath); // 先创建目录
        }

        string filePath = Path.Combine(dirPath, "CfgBase.cs");
        if (File.Exists(filePath)) return; // 已存在就跳过

        string content =
    @"using System.Collections;
using System.Collections.Generic;
using UnityEngine;

public class CfgBase
{
    public virtual void LoadData() { }
    public virtual void Release() { }
}";
        File.WriteAllText(filePath, content);
        Debug.Log("<color=#00ff00>CfgBase.cs Created!</color>");
    }


    static void CreateCfgMgr()
    {
        string dirPath = Path.Combine(DataPath);
        if (!Directory.Exists(dirPath))
        {
            Directory.CreateDirectory(dirPath); // 先创建目录
        }

        string filePath = Path.Combine(dirPath, "CfgMgr.cs");
        if (File.Exists(filePath)) return; // 已存在就跳过

        string content =
    @"using System.Collections;
using System.Collections.Generic;
using UnityEngine;

public class CfgMgr
{
    public static CfgMgr Instance { get; private set; } = new CfgMgr();

    public void Init()
    {
        // 初始化配置
    }
}";
        File.WriteAllText(filePath, content);
        Debug.Log("<color=#00ff00>CfgMgr.cs Created!</color>");
    }


    static void UpdateCfgMgr(List<string> classList)
    {
        string filePath = Path.Combine(BasePath, "CfgMgr.cs");
        if (!File.Exists(filePath))
        {
            Debug.Log("CfgMgr.cs已存在");
            return;
        }

        // 读取现有内容
        string content = File.ReadAllText(filePath);

        // 构建新的类成员
        string newFields = "";
        string newInits = "";
        foreach (string className in classList)
        {
            if (!content.Contains($"public {className} {className}")) // 避免重复添加
            {
                newFields += $"    public {className} {className} = new {className}();\n";
                newInits += $"        {className}.LoadData();\n";
            }
        }

        // 找到 Init 方法的插入点
        int initIndex = content.IndexOf("public void Init()");
        if (initIndex != -1)
        {
            int startIndex = content.IndexOf("{", initIndex) + 1;
            int endIndex = content.IndexOf("}", startIndex);

            if (startIndex != -1 && endIndex != -1)
            {
                // 插入初始化代码
                content = content.Substring(0, startIndex) + "\n" + newInits + content.Substring(endIndex);
            }
        }

        // 在类的最顶层插入字段
        int classStartIndex = content.IndexOf("{");
        if (classStartIndex != -1)
        {
            content = content.Insert(classStartIndex + 1, "\n" + newFields);
        }

        File.WriteAllText(filePath, content);
        Debug.Log("<color=#00ff00>CfgMgr.cs Updated!</color>");
    }



    /// <summary>
    /// 读取表格并保存脚本及json
    /// </summary>
    static void LoadData(string filePath, string fileName)
    {
        FileStream fileStream = null;
        DataSet result = null;

        // 打开文件流
        fileStream = File.Open(filePath, FileMode.Open, FileAccess.Read);

        // 创建 ExcelDataReader
        IExcelDataReader excelDataReader = ExcelReaderFactory.CreateOpenXmlReader(fileStream);

        result = excelDataReader.AsDataSet();
        // 生成 JSON 和模板
        CreateJson(result, fileName);
        CreateTemplate(result, fileName);

    }

    /// <summary>
    /// 生成json文件
    /// </summary>
    static void CreateJson(DataSet result, string fileName)
    {
        // 创建目录
        if (!Directory.Exists(jsonPath))
        {
            Directory.CreateDirectory(jsonPath);
        }

        // 获取表格有多少列 
        int columns = result.Tables[0].Columns.Count;
        // 获取表格有多少行 
        int rows = result.Tables[0].Rows.Count;

        TableData tempData;
        string value;
        JArray array = new JArray();

        //第一行为表头，第三行为类型 ，第二行为字段名 不读取
        for (int i = 3; i < rows; i++)
        {
            for (int j = 0; j < columns; j++)
            {
                // 获取表格中指定行指定列的数据 
                value = result.Tables[0].Rows[i][j].ToString();

                //if (string.IsNullOrEmpty(value))
                //{
                //    continue;
                //}
                tempData = new TableData();
                tempData.type = result.Tables[0].Rows[2][j].ToString();
                tempData.fieldName = result.Tables[0].Rows[1][j].ToString();
                tempData.value = value;

                dataList.Add(tempData);
            }
            if (dataList != null && dataList.Count > 0)
            {
                JObject tempo = new JObject();
                foreach (var item in dataList)
                {
                    try
                    {
                        //Debug.Log(item.type + " = " + item.value);
                        switch (item.type)
                        {
                            case "string":
                                if (item.value == null)
                                {
                                    tempo[item.fieldName] = "";
                                }
                                else
                                {
                                    tempo[item.fieldName] = ExcelExportTypeDefine.GetValue<string>(item.value);
                                }
                                break;
                            case "int":
                                if (string.IsNullOrEmpty(item.value))
                                {
                                    tempo[item.fieldName] = 0;
                                }
                                else
                                {
                                    tempo[item.fieldName] = ExcelExportTypeDefine.GetValue<int>(item.value);
                                }
                                break;
                            case "float":
                                if (string.IsNullOrEmpty(item.value))
                                {
                                    tempo[item.fieldName] = 0;
                                }
                                else
                                {
                                    tempo[item.fieldName] = ExcelExportTypeDefine.GetValue<float>(item.value);
                                }
                                break;
                            case "bool":
                                if (string.IsNullOrEmpty(item.value))
                                {
                                    tempo[item.fieldName] = false;
                                }
                                else
                                {
                                    tempo[item.fieldName] = ExcelExportTypeDefine.GetValue<bool>(int.Parse(item.value));
                                }
                                break;
                            case "string[]":
                                tempo[item.fieldName] = new JArray(ExcelExportTypeDefine.GetList<string>(item.value, ','));
                                break;
                            case "int[]":
                                tempo[item.fieldName] = new JArray(ExcelExportTypeDefine.GetList<int>(item.value, ','));
                                break;
                            case "float[]":
                                tempo[item.fieldName] = new JArray(ExcelExportTypeDefine.GetList<float>(item.value, ','));
                                break;
                            case "bool[]":
                                tempo[item.fieldName] = new JArray(ExcelExportTypeDefine.GetList<bool>(item.value, ','));
                                break;
                            case "EDailyTaskType":
                                tempo[item.fieldName] = new JArray(ExcelExportTypeDefine.GetList<int>(item.value, ','));
                                break;
                            case "ItemStruct":
                                tempo[item.fieldName] = JObject.Parse(ExcelExportTypeDefine.GetItemStructStr(item.value, false));
                                break;
                            case "ItemStruct[]":
                                tempo[item.fieldName] = JArray.Parse(ExcelExportTypeDefine.GetItemStructStr(item.value, true));
                                break;
                            case "KeyVal":
                                tempo[item.fieldName] = JObject.Parse(ExcelExportTypeDefine.GetKeyValStructStr(item.value, false));
                                break;
                            case "KeyVal[]":
                                tempo[item.fieldName] = JArray.Parse(ExcelExportTypeDefine.GetKeyValStructStr(item.value, true));
                                break;
                            case "LevelRange":
                                tempo[item.fieldName] = JObject.Parse(ExcelExportTypeDefine.GetLevelRangeStructStr(item.value, false));
                                break;
                        }
                    }
                    catch (Exception ex)
                    {
                        Debug.LogError("Exporting Config has Error: " + ex.Message + " The table is " + fileName + " ColumName" + item.fieldName + " Value" + item.value);
                    }
                }
                if (tempo != null)
                    array.Add(tempo);
                dataList.Clear();
            }
        }

        JObject o = new JObject();
        o["dataList"] = array;
        //o["version"] = "20200331";
        fileName = fileName.Replace(".xlsx", "");
        var jsonAddress = jsonPath + fileName + ".txt";
        if (File.Exists(jsonAddress))
        {
            File.Delete(jsonAddress);
        }

        string jsonStr = o.ToString();
        File.WriteAllText(jsonPath + fileName + ".txt", jsonStr);
        //if (Directory.Exists(ServerJsonPath))
        //{
        //    File.WriteAllText(ServerJsonPath + fileName + ".txt", jsonStr);
        //}
        //else
        //{
        //    Debug.LogError("ServerJsonPath not Exists! ignore file:" + ServerJsonPath + fileName + ".txt");
        //}
    }


    /// <summary>
    /// 生成实体类模板
    /// </summary>
    static void CreateTemplate(DataSet result, string fileName)
    {
        // 创建目录
        if (!Directory.Exists(scriptsPath))
        {
            Directory.CreateDirectory(scriptsPath);
        }

        field = "";
        for (int i = 0; i < result.Tables[0].Columns.Count; i++)
        {
            string typeStr = result.Tables[0].Rows[2][i].ToString();
            typeStr = ExcelExportTypeDefine.GetType(typeStr);
            if (typeStr.Contains("[]"))
            {
                typeStr = typeStr.Replace("[]", "");
                typeStr = string.Format("List<{0}>", typeStr);
            }

            string nameStr = result.Tables[0].Rows[1][i].ToString();
            if (string.IsNullOrEmpty(typeStr) || string.IsNullOrEmpty(nameStr)) continue;
            field += "public " + typeStr + " " + nameStr + (i + 1 == result.Tables[0].Columns.Count ? ";" : ";\n\t");
        }

        fileName = fileName.Replace(".xlsx", "");
        string tempStr = classTemp;
        tempStr = tempStr.Replace("@Name", fileName);
        tempStr = tempStr.Replace("@File1", field);
        tempStr = tempStr.Replace("@Date", DateTime.Now.ToLongDateString());

        var templatePath = scriptsPath + fileName + ".cs";
        if (File.Exists(templatePath))
        {
            File.Delete(templatePath);
        }

        File.WriteAllText(scriptsPath + fileName + ".cs", tempStr);
        //if (Directory.Exists(ServerScriptPath))
        //{
        //    File.WriteAllText(ServerScriptPath + fileName + ".cs", tempStr);
        //}
        //else
        //{
        //    Debug.LogError("ServerScriptPath not Exists! ignore file:" + ServerScriptPath + fileName + ".cs");
        //}

    }

    /// <summary>
    /// 字段
    /// </summary>
    static string field;

    private static string classTemp =
        "/*\r\n" +
        "Generated by: ScriptGenerator\r\n" +
        "Date: @Date\r\n" +
        "*/\r\n" +
        "using System;\r\n" +
        "using UnityEngine;\r\n" +
        "using System.Collections.Generic;\r\n\r\n" +
        "public partial class @Name : CfgBase\r\n" +
        "{\r\n    " +
        "public List<@NameCFG> dataList;\r\n\r\n    " +
        "public override void LoadData()\r\n    " +
        "{\r\n        " +
        "TextAsset ta = Resources.Load(\"GameRes/CFG/@Name\") as TextAsset;\r\n        " +
        "@Name data = JsonUtility.FromJson<@Name>(ta.text);\r\n        " +
        "dataList = data.dataList;\r\n    " + "" +
        "}\r\n\r\n    " + "" +
        "public override void Release()\r\n    " +
        "{\r\n        " + "" +
        "dataList = null;\r\n    " + "" +
        "}\r\n\r\n    " +
        "public @NameCFG Get(int id)\r\n    " +
        "{\r\n        " +
        "for (int i = 0; i < dataList.Count; i++)\r\n        " + "" +
        "{\r\n            " +
        "if (dataList[i].ID == id)\r\n            " +
        "{\r\n                " + "" +
        "return dataList[i];\r\n            " + "" +
        "}\r\n        " +
        "}\r\n        " +
        "return null;\r\n    " +
        "}\r\n}\r\n\r\n" +
        "[Serializable]\r\n" +
        "public class @NameCFG\r\n" +
        "{\r\n    " +
        "@File1" +
        "\r\n}";
}

public struct TableData
{
    public string fieldName;
    public string type;
    public string value;

    public override string ToString()
    {
        return string.Format("fieldName:{0} type:{1} value:{2}", fieldName, type, value);
    }
}


public class XFFile
{
    [UnityEditor.MenuItem("[GameConfig]/OpenConfigFolder", false, 2)]
    public static void OpenConfigFolder()
    {
        OpenFolder(ExcelExport.ConfigPath + "/");
    }

    [UnityEditor.MenuItem("[GameConfig]/DataFolder", false, 3)]
    public static void OpenGenFolder()
    {
        OpenFolder(ExcelExport.DataPath);
    }

    //[UnityEditor.MenuItem("[GameConfig]/Copy Scripts and Json Data From Unity Project To Server", false, 6)]
    //public static void CopyAllToServer()
    //{
    //    CopyDirectory(Application.dataPath + "/Hub/Resources/GameRes/CFG", ExcelExport.ServerJsonPath);
    //    CopyDirectory(Application.dataPath + "/Hub/Scripts/CFG", ExcelExport.ServerTopScriptPath);
    //}

    public static void OpenFolder(string filePath)
    {
        var folderPath = Path.GetDirectoryName(filePath);
        // 检查文件夹路径是否存在
        if (Directory.Exists(folderPath))
        {
            // 根据当前平台选择打开文件资源管理器的命令
            string explorerCommand = "";
            if (Application.platform == RuntimePlatform.WindowsEditor || Application.platform == RuntimePlatform.WindowsPlayer)
            {
                explorerCommand = "explorer";
            }
            else if (Application.platform == RuntimePlatform.OSXEditor || Application.platform == RuntimePlatform.OSXPlayer)
            {
                explorerCommand = "open";
                // 添加 -R 参数以打开文件夹
                folderPath = "-R \"" + folderPath + "\"";
            }
            else
            {
                Debug.LogWarning("该平台不支持打开文件夹");
                return;
            }

            // 打开文件资源管理器并指定文件夹路径
            System.Diagnostics.Process.Start(explorerCommand, folderPath);
        }
        else
        {
            Debug.LogError("文件夹路径不存在：" + folderPath);
        }
    }

    public static void CopyDirectory(string sourceDir, string targetDir, bool ignoreMeta = true)
    {
        // 创建目标目录
        if (!Directory.Exists(targetDir))
        {
            Directory.CreateDirectory(targetDir);
        }

        if (!Directory.Exists(sourceDir))
        {
            Debug.LogError("Source Directory not Exists! " + sourceDir);
            return;
        }

        // 复制所有文件
        foreach (string file in Directory.GetFiles(sourceDir))
        {
            string fileName = Path.GetFileName(file);
            if (ignoreMeta && fileName.EndsWith(".meta")) { continue; }
            string destFile = Path.Combine(targetDir, fileName);
            File.Copy(file, destFile, true); // true 表示如果文件已存在则覆盖
            Debug.LogWarning("Copy File:[" + file + "] -> [" + destFile + "]");
        }

        // 递归复制所有子目录
        foreach (string directory in Directory.GetDirectories(sourceDir))
        {
            string dirName = Path.GetFileName(directory);
            string destDir = Path.Combine(targetDir, dirName);
            CopyDirectory(directory, destDir);
        }
    }
}