using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using UnityEngine;
using ExcelDataReader;
using Newtonsoft.Json;
using UnityEditor;

public class FileExporter
{
    private readonly List<FileInfo> srcFiles;
    private readonly List<bool> isUseFiles;
    private readonly string savePath;
    private readonly int sheetCount;
    private readonly int headerRows;

    public FileExporter(List<FileInfo> _srcFiles, List<bool> _isUseFiles, string _savePath, int _sheetCount,
        int _headerRows)
    {
        srcFiles = _srcFiles;
        isUseFiles = _isUseFiles;
        savePath = _savePath;
        sheetCount = _sheetCount;
        headerRows = _headerRows;
    }

    public int SaveJsonFiles()
    {
        return ReadAllTables(SaveSheetJson);
    }

    public int SaveCsFiles()
    {
        return ReadAllTables(SaveSheetCs);
    }

    public int SaveSOFiles()
    {
        return ReadAllTables(SaveTableAsset);
    }
    
    
    #region Read Table

    int ReadAllTables(Func<DataTable, string, int> exportFunc)
    {
        if (srcFiles == null || srcFiles.Count <= 0)
        {
            Debug.LogError("Error! No Excel Files!");
            return -1;
        }

        int result = 0;
        for (var i = 0; i < srcFiles.Count; i++)
        {
            if (isUseFiles[i])
            {
                var file = srcFiles[i];
                result += ReadTable(file.FullName, FileNameNoExt(file.Name), exportFunc);
            }
        }

        return result;
    }

    int ReadTable(string path, string fileName, Func<DataTable, string, int> exportFunc)
    {
        int result = 0;
        using (var stream = File.Open(path, FileMode.Open, FileAccess.Read)) 
        {
            using (var reader = ExcelReaderFactory.CreateReader(stream))
            {
                int tableSheetNum = reader.ResultsCount;
                if (tableSheetNum < 1)
                {
                    Debug.LogError("Excel file is empty: " + path);
                    return -1;
                }
                
                var dataSet = reader.AsDataSet();
                
                int checkCount = sheetCount <= 0 ? tableSheetNum : sheetCount;
                for (int i = 0; i < checkCount; i++)
                {
                    if (i < tableSheetNum)
                    {
                        string name = checkCount == 1 ? 
                            fileName :
                            fileName + "_" + dataSet.Tables[i].TableName;
                        //result += SaveJson(dataSet.Tables[i], name);
                        result += exportFunc(dataSet.Tables[i], name);
                    }
                }
            }
        }
        return result;
    }

    #endregion
    
    
    #region Save Json Files
    
    int SaveSheetJson(DataTable sheet, string fileName)
    {
        if (sheet.Rows.Count <= 0)
        {
            Debug.LogError("Excel Sheet is empty: " + sheet.TableName);
            return -1;
        }
        
        int columns = sheet.Columns.Count;
        int rows = sheet.Rows.Count;

        List<Dictionary<string, object>> tData = new List<Dictionary<string, object>>();
                
        for (int i = headerRows; i < rows; i++)
        {
            Dictionary<string, object> rowData = new Dictionary<string, object>();
            for (int j = 0; j < columns; j++)
            {
                string key = sheet.Rows[0][j].ToString();
                rowData[key] = sheet.Rows[i][j];
            }

            tData.Add(rowData);
        }

        string json = JsonConvert.SerializeObject(tData, Formatting.Indented);
        
        // save to file
        string dstFolder = savePath + "/TableJson/";
        if (!Directory.Exists(dstFolder))
        {
            Directory.CreateDirectory(dstFolder);
        }
        
        string path = dstFolder + fileName + ".json";
        using (FileStream fileStream = new FileStream(path, FileMode.Create, FileAccess.Write))
        {
            using (TextWriter textWriter = new StreamWriter(fileStream, Encoding.UTF8))
            {
                textWriter.Write(json);
                Debug.Log("File saved: " + path);
                return 1;
            }
        }
    }

    #endregion


    
    #region Save Class File

    int SaveSheetCs(DataTable sheet, string fileName)
    {
        if (sheet.Rows.Count < 2)
        {
            Debug.LogError("Excel Sheet is empty: " + sheet.TableName);
            return -1;
        }

        // write table data class
        int r1 = SaveTableClass(sheet, fileName);
        
        // write scriptable object TableSet class
        if (r1 > 0)
        {
            int r2 = SaveTableSet(sheet, fileName);
            return r2;
        }
        
        return -1;
    }

    int SaveTableClass(DataTable sheet, string fileName)
    {
        StringBuilder sb = new StringBuilder();
        sb.AppendLine("/// <summary>");
        sb.AppendLine($"/// Auto Generated Class by ExcelTool. From Table: {fileName}");
        sb.AppendLine("/// </summary>");

        string keyType = sheet.Rows[1][0].ToString();
        fileName = "Table_" + fileName;
        
        sb.AppendFormat("public class {0} : TableBase<{1}>\r\n{{", fileName, keyType);
        sb.AppendLine();

        foreach (DataColumn column in sheet.Columns)
        {
            string name = sheet.Rows[0][column].ToString();
            string type = sheet.Rows[1][column].ToString();
            string comment = sheet.Rows[2][column].ToString();
            
            if (string.IsNullOrEmpty(name)) continue;
            if (string.IsNullOrEmpty(type)) type = "string";
            if (!string.IsNullOrEmpty(comment))
            {
                comment = comment.Replace("\n", ",").Replace("\r", ",");
            }
            
            sb.AppendFormat("\tpublic {0} {1}; // {2}", type, name, comment);
            sb.AppendLine();
        }
        
        sb.Append('}');
        sb.AppendLine();
        
        // save to file
        string dstFolder = savePath + "/TableClass/";
        if (!Directory.Exists(dstFolder))
        {
            Directory.CreateDirectory(dstFolder);
        }
        
        string path = dstFolder + fileName + ".cs";
        using (FileStream fileStream = new FileStream(path, FileMode.Create, FileAccess.Write))
        {
            using (TextWriter textWriter = new StreamWriter(fileStream, Encoding.UTF8))
            {
                textWriter.Write(sb.ToString());
                Debug.Log("Table saved: " + path);
                return 1;
            }
        }
    }
    
    
    int SaveTableSet(DataTable sheet, string fileName)
    {
        StringBuilder sb = new StringBuilder();
        sb.AppendLine("/// <summary>");
        sb.AppendLine($"/// Auto Generated ScriptableObject TableSet.");
        sb.AppendLine("/// </summary>");

        string keyType = sheet.Rows[1][0].ToString();
        string tSetName = "TSet_" + fileName;
        string tClassName = "Table_" + fileName;
        
        sb.AppendFormat("public class {0} : TableSet<{1}, {2}>\r\n{{", tSetName, keyType, tClassName);
        sb.AppendLine();
        sb.Append('}');
        sb.AppendLine();
        
        // save to file
        string dstFolder = savePath + "/TableClass/";
        if (!Directory.Exists(dstFolder))
        {
            Directory.CreateDirectory(dstFolder);
        }
        
        string path = dstFolder + tSetName + ".cs";
        using (FileStream fileStream = new FileStream(path, FileMode.Create, FileAccess.Write))
        {
            using (TextWriter textWriter = new StreamWriter(fileStream, Encoding.UTF8))
            {
                textWriter.Write(sb.ToString());
                Debug.Log("Table Set saved: " + path);
                return 1;
            }
        }
    }

    #endregion


    #region Save ScriptableObject

    int SaveTableAsset(DataTable sheet,  string fileName)
    {
        if (sheet.Rows.Count < 2)
        {
            Debug.LogError("Excel Sheet is empty: " + sheet.TableName);
            return -1;
        }
        
        string tClassName = "Table_" + fileName;
        string tSetName = "TSet_" + fileName;

        // create Main Asset of table Set type by reflection
        var methods = typeof(ScriptableObject).GetMethods().Where(m => m.Name == "CreateInstance");
        var methodInfo = methods.First(m => m.IsGenericMethod);
        
        Assembly assembly = Assembly.Load("Assembly-CSharp");
        var tSetType = assembly.GetType(tSetName);
        if (tSetType == null) return -1;
        MethodInfo createMethod = methodInfo.MakeGenericMethod(tSetType);

        UnityEngine.Object assetObj = (UnityEngine.Object) Activator.CreateInstance(tSetType);
        createMethod.Invoke(assetObj, null);

        // save to file
        string dstFolder = savePath + "/TableAssets/";
        if (!Directory.Exists(dstFolder))
        {
            Directory.CreateDirectory(dstFolder);
        }

        int idx = dstFolder.IndexOf("Assets/", StringComparison.Ordinal);
        string path = dstFolder.Substring(idx) + tSetName + ".asset";
        AssetDatabase.CreateAsset(assetObj, path);
        Debug.Log("Main Asset saved: " + path);
        
        // create asset of every row data, and add them to main asset
        for (int i = headerRows; i < sheet.Rows.Count; i++)
        {
            var tableType = assembly.GetType(tClassName);
            
            UnityEngine.Object dataObj = (UnityEngine.Object) Activator.CreateInstance(tableType);
            createMethod.Invoke(dataObj, null);

            // set every field data
            object keyVal = null;
            for (var j = 0; j < sheet.Columns.Count; j++)
            {
                string name = sheet.Rows[0][j].ToString();
                string type = sheet.Rows[1][j].ToString();
                string val = sheet.Rows[i][j].ToString();

                if (string.IsNullOrEmpty(name)) continue;
                if (string.IsNullOrEmpty(type)) type = "string";
                
                object o = SetObjectFiled(dataObj, tableType.GetField(name), type, val);
                if (j == 0)
                {
                    keyVal = o;
                }
            }
            
            // set key data with first column value
            tableType.GetField("tKey").SetValue(dataObj, keyVal);
            
            // add data asset to main asset
            var methodAdd = tSetType.GetMethod("Add");
            methodAdd.Invoke(assetObj, new[] {keyVal, dataObj});

            dataObj.name = tClassName + "_" + keyVal;
            AssetDatabase.AddObjectToAsset(dataObj, path);
        }

        EditorUtility.SetDirty(assetObj);
        AssetDatabase.SaveAssets();
        AssetDatabase.Refresh();

        return 1;
    }

  
    object SetObjectFiled(object obj, FieldInfo field, string type, string param)
    {
        object pObj = param;
        switch (type.ToLower())
        {
            case "string":
                field.SetValue(obj, param);
                break;
            case "bool":
                pObj = bool.Parse(param);
                field.SetValue(obj, bool.Parse(param));
                break;
            case "byte":
                pObj = byte.Parse(param);
                field.SetValue(obj, byte.Parse(param));
                break;
            case "int":
                pObj = int.Parse(param);
                field.SetValue(obj, int.Parse(param));
                break;
            case "short":
                pObj = short.Parse(param);
                field.SetValue(obj, short.Parse(param));
                break;
            case "long":
                pObj = long.Parse(param);
                field.SetValue(obj, long.Parse(param));
                break;
            case "float":
                pObj = float.Parse(param);
                field.SetValue(obj, float.Parse(param));
                break;
            case "double":
                pObj = double.Parse(param);
                field.SetValue(obj, double.Parse(param));
                break;
            case "decimal":
                pObj = decimal.Parse(param);
                field.SetValue(obj, decimal.Parse(param));
                break;
            default:
                Assembly assembly = Assembly.Load("Assembly-CSharp");
                var t = assembly.GetType(type);
                if (t != null)
                {
                    if (t.IsEnum)
                    {
                        pObj = Enum.Parse(t, param);
                        field.SetValue(obj, Enum.Parse(t, param));
                    }
                }
                else
                {
                    field.SetValue(obj, param);  
                }
                break;
        }

        return pObj;
    }
    
    #endregion
    
    
    
    
    
    string FileNameNoExt(string filename)
    {
        int length;
        if ((length = filename.LastIndexOf('.')) == -1)
            return filename;
        return filename.Substring(0, length);
    }

}
