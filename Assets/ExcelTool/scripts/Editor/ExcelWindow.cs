using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using UnityEditor;
using UnityEngine;

[Serializable]
public class ExcelWindow : EditorWindow
{
    private const string KeyFilePath = "Excel_FilePath";
    private const string KeyOutPath = "Excel_OutPath";
    private const string KeySheetCount = "Excel_SheetCount";
    private const string KeyIncSub = "Excel_IncSub";

    private string filePath; // source excel file path
    private string outPath; // output path
    private bool isIncSub; // whether include subdirectories
    private int sheetCount = 1; // load sheet count. 0 means all.
    
    private int excelFileCount;
    private List<FileInfo> excelFiles; 
    private List<bool> isUseFiles;

    private const int H1 = 140;
    private const int H2 = 160;
    private const int H3 = 90;

    private int genResult;
    
    enum LogState
    {
        None,
        NoSrcFolder,
        NoOutFolder,
        NoExcelFiles,
        WaitngOperate,
        GenSuccess,
        GenFailed,
        NameDuplicated
    }
    private static LogState curLogState;
    
    
    [MenuItem("Window/Excel Tool Window")]
    public static void ShowWindow()
    {
        //GetWindowWithRect(typeof(ExcelWindow), new Rect(50, 100, 600, 600));
        GetWindow(typeof(ExcelWindow), false, "Excel Tool");
    }

    private void OnEnable()
    {
        filePath = EditorPrefs.GetString(KeyFilePath, Application.dataPath);
        outPath = EditorPrefs.GetString(KeyOutPath, Application.dataPath + "/Tables");
        sheetCount = EditorPrefs.GetInt(KeySheetCount, 1);
        isIncSub = EditorPrefs.GetBool(KeyIncSub, false);

        LoadExcelFiles();
    }


    void OnGUI()
    {
        DrawPath();
        
        DrawFileInfos();

        DrawExcelFiles();
        
        DrawInfo();
        
        DrawLogArea();
    }

    /// <summary>
    /// browse excel file path and output file path
    /// </summary>
    void DrawPath()
    {
        EditorGUILayout.LabelField("Settings", EditorStyles.boldLabel);

        GUIStyle tfstyle = new GUIStyle(GUI.skin.textField){alignment = TextAnchor.MiddleLeft, fixedHeight = 19}; 
        
        // choose excel path
        GUILayout.Space(5);
        EditorGUILayout.BeginHorizontal();
        EditorGUIUtility.labelWidth = 80;
        filePath = EditorGUILayout.TextField ("Excel Path:", filePath, tfstyle);
        
        if (GUILayout.Button("Browse", GUILayout.Width(80)))
        {
            string tmpPath = EditorUtility.OpenFolderPanel("Choose Excel Files Forlder", "", "");
            if (!string.IsNullOrEmpty(tmpPath) && tmpPath != filePath)
            {
                filePath = tmpPath;
                LoadExcelFiles();
            }
        }
        EditorGUILayout.EndHorizontal();
        
        // choose output path
        GUILayout.Space(5);
        EditorGUILayout.BeginHorizontal();
        EditorGUIUtility.labelWidth = 80;
        outPath = EditorGUILayout.TextField ("Output Path:", outPath, tfstyle);
        
        if (GUILayout.Button("Browse", GUILayout.Width(80)))
        {
            outPath = EditorUtility.OpenFolderPanel("Choose Excel Files Forlder", "", "");
            EditorPrefs.SetString(KeyOutPath, outPath);
        }
        EditorGUILayout.EndHorizontal();
    }

   
    /// <summary>
    /// file count, checkbox of whether scan subdir, and refresh button
    /// </summary>
    void DrawFileInfos()
    {
        GUILayout.Space(20);
        EditorGUILayout.LabelField("Excel Files", EditorStyles.boldLabel, GUILayout.Width(100));

        GUILayout.Space(5);
        EditorGUILayout.BeginHorizontal();
        EditorGUILayout.LabelField("File count: " + excelFileCount, GUILayout.Width(85));
        bool _isIncSub = GUILayout.Toggle(isIncSub, "Include subdirectories");

        if (_isIncSub != isIncSub)
        {
            isIncSub = _isIncSub;
            EditorPrefs.SetBool(KeyIncSub, isIncSub);
            LoadExcelFiles();
        }
        
        GUILayout.Space(20);
        if (GUILayout.Button("Refresh", GUILayout.Width(80)))
        {
            LoadExcelFiles();
        }
        GUILayout.FlexibleSpace();
        EditorGUILayout.EndHorizontal();
    }

    /// <summary>
    /// list of excle tables
    /// </summary>
    void DrawExcelFiles()
    {
        Rect rect = new Rect(5, H1, position.width - 10, position.height - H1 - H2 - H3 - 30);
        EditorGUI.DrawRect(rect, new Color(0.8f, 0.8f, 0.8f));
        GUILayout.BeginArea(new Rect(rect.x + 5, rect.y + 5, rect.width - 10, rect.height - 10));

        if (string.IsNullOrEmpty(filePath))
        {
            curLogState = LogState.NoSrcFolder; 
        }
        else if(excelFiles == null || excelFiles.Count < 1)
        {
            curLogState = LogState.NoExcelFiles;
        }
        else
        {
            for (int i = 0; i < excelFileCount; i++)
            {
                isUseFiles[i] = GUILayout.Toggle(isUseFiles[i], excelFiles[i].FullName);
            }

            if (string.IsNullOrEmpty(outPath))
            {
                curLogState = LogState.NoOutFolder; 
            }
            else if (curLogState != LogState.GenSuccess && curLogState != LogState.GenFailed && curLogState != LogState.NameDuplicated)
            {
                curLogState = LogState.WaitngOperate;
            }
        }
        
        GUILayout.EndArea();
    }

    void DrawInfo()
    {
        Rect rect = new Rect(5, position.height - H2 - H3 - 20, position.width - 10, H2);
        EditorGUI.DrawRect(rect, new Color(0.7f, 0.7f, 0.7f));
        GUILayout.BeginArea(rect);
        
        EditorGUILayout.LabelField("Notes:", EditorStyles.boldLabel);
        EditorGUILayout.LabelField("1. First row of table: name of data (表格第一行：数据名称)", new GUIStyle {fontSize = 12});
        EditorGUILayout.LabelField("2. Second row of table: type of data, only support base type and Enum \n(表格第二行：数据类型, 仅支持基本类型和Enum)", new GUIStyle {fontSize = 12}, GUILayout.Height(30));
        EditorGUILayout.LabelField("3. Third row of table: comment of data, could be empty (表格第三行：注释，可以为空)", new GUIStyle {fontSize = 12});
        EditorGUILayout.LabelField("4. First column data will be the key to find table (第一列数据将被用作查找表的Key)", new GUIStyle {fontSize = 12});
        EditorGUILayout.LabelField("5. All Excel filenames cannot be duplicated  (表格名不能重复)", new GUIStyle {fontSize = 12});

        EditorGUILayout.LabelField("<color=#B22222>6. If you want to generate Asset, you must have auot generated class first! \n(如果要生成SO资源文件，必须先有自动生成的类文件)</color>", new GUIStyle {richText = true, fontSize = 12}, GUILayout.Height(40));
        
        GUILayout.EndArea();
    }
    

    void DrawLogArea()
    {
        //EditorGUI.DrawRect(new Rect(5, position.height - H3 - 10, position.width - 10, H3), Color.green);
        GUILayout.BeginArea(new Rect(5, position.height - H3 - 10, position.width - 10, H3));

        GUILayout.Space(5);
        
        GUILayout.BeginHorizontal();
        EditorGUILayout.LabelField("Sheet Count (0 means all): ", GUILayout.Width(170));
        string sheetNum = EditorGUILayout.TextField("", sheetCount.ToString(), GUILayout.Width(40));
        int.TryParse(sheetNum, out sheetCount);
        EditorPrefs.SetInt(KeySheetCount, sheetCount);
        GUILayout.FlexibleSpace();
        GUILayout.EndHorizontal();
        
        GUILayout.Space(5);
        
        EditorGUILayout.BeginHorizontal();
        GUI.backgroundColor = new Color(0.85f, 1f, 1f);
        
        if (GUILayout.Button("Json", GUILayout.Width(80), GUILayout.Height(30)))
        {
            GenerateJson();
        }
        GUILayout.Space(10);
        
        GUI.backgroundColor = new Color(0.7f, 1f, 0.7f);
        if (GUILayout.Button("Class", GUILayout.Width(80), GUILayout.Height(30)))
        {
            GenerateClass();
        }
        GUILayout.Space(10);
        
        if (GUILayout.Button("Asset", GUILayout.Width(80), GUILayout.Height(30)))
        {
            GenerateAsset();
        }
        
        EditorGUILayout.EndHorizontal();
    
        GUILayout.Space(10);
        
        string logInfo = "";
        switch (curLogState)
        {
            case LogState.None:
                break;
            case LogState.NoSrcFolder:
                logInfo = "<color=red>Please select a Folder of excel files!</color>";
                break;
            case LogState.NoOutFolder:
                logInfo = "<color=red>Please select an output Folder!</color>";
                break;
            case LogState.NoExcelFiles:
                logInfo = "<color=red>No Excel Files! Click refresh button to load files.</color>";
                break;
            case LogState.WaitngOperate:
                logInfo = "Click Button to Start";
                break;
            case LogState.GenSuccess:
                logInfo = $"Sucdcess! Generated {genResult} Files.";
                break;
            case LogState.GenFailed:
                logInfo = "<color=red>Generate Failed!</color>";
                break;
            case LogState.NameDuplicated:
                logInfo = "<color=red>Error: Name Duplicated!</color>";
                break;
        }
        
        EditorGUILayout.LabelField(logInfo, new GUIStyle(){richText = true});
        GUILayout.EndArea();
    }
    
    void LoadExcelFiles()
    {
        if(string.IsNullOrEmpty(filePath)) return;
  
        EditorPrefs.SetString(KeyFilePath, filePath);
        
        excelFiles = new List<FileInfo>();
        GetFolderFiles(excelFiles, filePath, isIncSub);
        excelFileCount = excelFiles.Count;

        isUseFiles = new List<bool>(Enumerable.Repeat(true, excelFileCount));
        curLogState = HaveDuplicatedName() ? LogState.NameDuplicated : LogState.None;
        genResult = 0;
    }

    void GetFolderFiles(List<FileInfo> files, string path, bool incSub = false)
    {
        DirectoryInfo folder = new DirectoryInfo(path);
        if (!folder.Exists) return;

        files.AddRange(GetFilesByExtensions(folder, ".xlsx", ".xls"));

        if (incSub)
        {
            foreach (var info in folder.GetDirectories())
            {
                GetFolderFiles(files, info.FullName, true);
            }
        }
    }
    
    
    IEnumerable<FileInfo> GetFilesByExtensions(DirectoryInfo dir, params string[] extensions)
    {
        if (extensions == null) return null;
        
        IEnumerable<FileInfo> files = dir.EnumerateFiles();
        return files.Where(f => extensions.Contains(f.Extension.ToLower()));
    }


    bool HaveDuplicatedName()
    {
        if (excelFiles == null || excelFiles.Count <= 0) return false;
        
        List<string> names = new List<string>();
        foreach (var file in excelFiles)
        {
            string fileName = Path.GetFileNameWithoutExtension(file.FullName);
            if (names.Contains(fileName))
            {
                return true;
            }
            names.Add(fileName);
        }
        return false;
    }

    bool CanGenerate()
    {
        if (curLogState == LogState.NameDuplicated || curLogState == LogState.NoExcelFiles ||
            curLogState == LogState.NoOutFolder || curLogState == LogState.NoSrcFolder) return false;
        if (excelFiles == null || excelFiles.Count <= 0)
        {
            Debug.LogError("Error! No Excel Files!");
            return false;
        }
        
        if (string.IsNullOrEmpty(outPath))
        {
            Debug.LogError("Error! No Output Folder!");
            return false;
        }

        return true;
    }

    void GenerateJson()
    {
        if(!CanGenerate()) return;
        
        FileExporter exporter = new FileExporter(excelFiles, isUseFiles, outPath, sheetCount, 3);
        genResult = exporter.SaveJsonFiles();
        curLogState = genResult > 0 ? LogState.GenSuccess : LogState.GenFailed;
        AssetDatabase.Refresh();
    }

    void GenerateClass()
    {
        if(!CanGenerate()) return;

        FileExporter exporter = new FileExporter(excelFiles, isUseFiles, outPath, sheetCount, 3);
        genResult = exporter.SaveCsFiles();
        curLogState = genResult > 0 ? LogState.GenSuccess : LogState.GenFailed;
        AssetDatabase.Refresh();
    }

    void GenerateAsset()
    {
        if(!CanGenerate()) return;
        
        FileExporter exporter = new FileExporter(excelFiles, isUseFiles, outPath, sheetCount, 3);
        genResult = exporter.SaveSOFiles();
        curLogState = genResult > 0 ? LogState.GenSuccess : LogState.GenFailed;
        AssetDatabase.Refresh();
    }

    
    void ClearDir(string path)
    {
        if (!Directory.Exists(path))
        {
            Directory.CreateDirectory(path);
        }

        DirectoryInfo dir = new DirectoryInfo(path);
        FileInfo[] files = dir.GetFiles();

        foreach (var item in files)
        {
            File.Delete(item.FullName);
        }
        if (dir.GetDirectories().Length != 0)
        {
            foreach (var item in dir.GetDirectories())
            {
                if (!item.ToString().Contains("$") && (!item.ToString().Contains("Boot")))
                {
                    ClearDir(Path.Combine(dir.ToString(), item.ToString()));
                }
            }
        }
    }

    
}
