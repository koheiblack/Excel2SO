using UnityEngine;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using UnityEditor;
using System.Xml.Serialization;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using System.Text;
using System.Reflection;

public class ExcelScriptableObjectImporter : AssetPostprocessor
{
    static readonly string[] extXlsFilters = new string[] { ".xls", ".xlsx" };
    static bool hashCheck = true;
//    static readonly string extCsFilter = ".cs";

    static List<string> GenerateClassFileNames = new List<string>();
    static List<string> ImportFileNames = new List<string>();

    [MenuItem("Assets/ExcelImporter/Reimport Excel Files")]
    static void ReimportExcelFiles()
    {
        hashCheck = false;
        AssetDatabase.ImportAsset("Assets", ImportAssetOptions.ForceUpdate | ImportAssetOptions.DontDownloadFromCacheServer | ImportAssetOptions.ImportRecursive);
    }

	//[MenuItem("Assets/ExcelImporter/Force Reimport")]
	static void ClearUserData (MenuCommand command) {
		string str = "";

		foreach (var obj in Selection.objects) {
			string selectionPath = AssetDatabase.GetAssetPath(obj);

            AssetImporter assetImporter = AssetImporter.GetAtPath(selectionPath);
            string userdatastring = assetImporter != null ? assetImporter.userData : string.Empty;
			str += "[" + selectionPath + ":" + userdatastring + "]\n";
			assetImporter.userData = ""; // ハッシュを強制クリア
		}

		Debug.Log(str);
        AssetDatabase.Refresh(ImportAssetOptions.ForceUpdate);

        // ここでリインポートも行う
        foreach (var obj in Selection.objects) {
			string selectionPath = AssetDatabase.GetAssetPath(obj);
			AssetDatabase.ImportAsset( selectionPath/*, ImportAssetOptions.ForceUpdate*/ );
		}
	}

    static void AddGenerateClassFileName(string name)
    {
        if ( !GenerateClassFileNames.Contains(name) )
        {
            GenerateClassFileNames.Add(name);
        }
    }

    static void AddImportFileNames(string name)
    {
        if ( !ImportFileNames.Contains(name) )
        {
            ImportFileNames.Add(name);
        }
    }


    static void OnPostprocessAllAssets(string[] importedAssets, string[] deletedAssets, string[] movedAssets, string[] movedFromAssetPaths)
    {
        foreach ( string asset in importedAssets )
        {
            //Debug.Log("Import: " + asset + " " + Path.GetExtension(asset));
            string ext = Path.GetExtension(asset);
            bool isXls = false;
            foreach ( string extfilter in extXlsFilters )
            {
                if ( ext.Equals(extfilter) )
                {
                    isXls = true;
                    break;
                }
            }
            if ( !isXls ) continue;

            AssetImporter assetImporter = AssetImporter.GetAtPath(asset);
            string userdatastring = assetImporter != null ? assetImporter.userData : string.Empty;
            Importer importer = new Importer();
            if ( assetImporter == null || !importer.IsEqualHash(asset, userdatastring) || !hashCheck )
            {
                importer.CreateClasses(asset);
                // ToDo: wait compiling
                if ( GenerateClassFileNames.Count > 0 )
                {
                    foreach ( string fn in GenerateClassFileNames )
                    {
                        AssetDatabase.ImportAsset(fn, ImportAssetOptions.ForceSynchronousImport);
                    }
                    GenerateClassFileNames.Clear();
                    AssetDatabase.ImportAsset(asset);
                }
                else if ( importer.CreateAssets(asset) )
                {
                    if ( assetImporter != null ) assetImporter.userData = importer.xlsHash;
                    AssetDatabase.Refresh(ImportAssetOptions.ForceUpdate);
                }
                else
                {
                    Debug.LogError("Import failed: " + asset);
                }
            }
        }
        hashCheck = true;
    }

 
    public class Importer
    {
        string supportedFormatHeader = "#XLS2SO";
        string classExportPath = "Assets/ExcelImporter/Classes/";
        string exportDirectory = string.Empty;
        string exportDirectoryReplaceFrom = "Assets";
        string exportDirectoryReplaceTo = "Assets/Resources";
        string fileName = string.Empty;
        string filePath = string.Empty;

        public enum OutputBookType
        {
            Separate,
            Array,
            List,
            Dictionary,
            Individual,
        }
        OutputBookType outputBookType = OutputBookType.Separate;
        public enum OutputSheetType
        {
            Array,
            List,
        }
        OutputSheetType outputSheetType = OutputSheetType.List;
        List<ExcelColParameter> typeList = new List<ExcelColParameter>();
        string className = string.Empty;
        string subClassName = string.Empty;
        string paramStructName = "Param";
        string paramListFieldName = "param";
        string sheetStructName = "Sheet";
        string sheetListFieldName = "sheet";
        string sheetStructNameFieldName = "name";
        string sheetStructListFieldName = "list";
        string structTagName = "tag";
        string subClassTagName = "";
        string subClassFieldNameSuffix = "List";
        List<string> outputClassList = new List<string>();
        int startRowNumber = 0;
        int endRowNumber = 0;
        int fieldRowNumber = 0;
        int typeRowNumber = 0;
        int defaultValueRowNumber = 0;
        bool isExternalParam = false;
        bool isIgnoredEmptyElement = true;
        public string xlsHash = "xls";


        enum ValueType
        {
            BOOL,
            STRING,
            INT,
            FLOAT,
            DOUBLE,
            VECTOR2,
            VECTOR3,
            COLOR,
            HASH,

            ENUM,
            STRUCT,
            CLASS,
            UNKNOWN,
        }

        string[] valueTypeNames = new string[] {
            "bool",
            "string",
            "int",
            "float",
            "double",
            "Vector2",
            "Vector3",
            "Color",
            "int",

            "",
            "",
            "",
            "",
        };

        string[] csTypeNames = new string[] {
            "System.Boolean",
            "System.String",
            "System.Int32",
            "System.Single",
            "System.Double",
            "UnityEngine.Vector2",
            "UnityEngine.Vector3",
            "UnityEngine.Color",
            "System.Int32",

            "",
            "",
            "",
            "",
        };

        class ExcelColParameter
        {
            public ValueType type;
            public string typeName;
            public string cstypeName;
            public string name;
            public string structName;
            public bool isEnable = false;
            public bool isArray = false;
            public bool isNextArray = false;
            public bool isStructArray = false;
            public bool isNextStruct = false;
            public int structNestCount = 0;
            public int colNumber;

            public void CopyFrom(ExcelColParameter from)
            {
                type = from.type;
                typeName = from.typeName;
                cstypeName = from.cstypeName;
                name = from.name;
                structName = from.structName;
                isEnable = from.isEnable;
                isArray = from.isArray;
                isNextArray = from.isNextArray;
                isStructArray = from.isStructArray;
                isNextStruct = from.isNextStruct;
                structNestCount = from.structNestCount;
                colNumber = from.colNumber;
            }
        }

        public bool IsEqualHash(string filename, string hash)
        {
            using ( FileStream stream = File.Open(filename, FileMode.Open, FileAccess.Read) )
            {
                System.Security.Cryptography.MD5CryptoServiceProvider md5 = new System.Security.Cryptography.MD5CryptoServiceProvider();
                xlsHash = BitConverter.ToString(md5.ComputeHash(stream));
            }
            if ( xlsHash.Equals(hash) )
            {
                return true;
            }

            return false;
        }

        IWorkbook ReadXlsFile(FileStream stream, string filename)
        {
            IWorkbook book;
            if ( Path.GetExtension(filename).Equals(".xls") )
            {
                book = new HSSFWorkbook(stream);
            }
            else
            {
                book = new XSSFWorkbook(stream);
            }

            return book;
        }

        public bool CreateClasses(string filename)
        {
            filePath = filename;
            exportDirectory = Path.GetDirectoryName(filePath).Replace(exportDirectoryReplaceFrom, exportDirectoryReplaceTo);
            fileName = Path.GetFileNameWithoutExtension(filePath);
            className = fileName;

            using ( FileStream stream = File.Open(filename, FileMode.Open, FileAccess.Read) )
            {
                IWorkbook book = ReadXlsFile(stream, filename);

                if ( !IsSupportedFormat(book) ) return true;

                if ( outputBookType == OutputBookType.Individual )
                {
                    ExportClassDefinitionIndividual(book);
                }
                else if ( outputBookType != OutputBookType.Separate )
                {
                    for ( int si = 0 ; si < book.NumberOfSheets ; si++ )
                    {
                        ISheet sheet = book.GetSheetAt(si);
                        if ( !IsValidSheet(sheet) ) continue;    // comment sheet

                        //LoadSettings(sheet);
                        if ( GenerateClassParameter(sheet) )
                        {
                            ExportClassDefinition(sheet);
                        }
                        break;
                    }
                }
                else
                {
                    for ( int si = 0 ; si < book.NumberOfSheets ; si++ )
                    {
                        ISheet sheet = book.GetSheetAt(si);
                        if ( !IsValidSheet(sheet) ) continue;    // comment sheet

                        LoadSettings(sheet);
                        if ( GenerateClassParameter(sheet) )
                        {
                            ExportClassDefinition(sheet);
                        }
                    }
                }
            }
            return true;
        }

        public bool CreateAssets(string filename)
        {
            filePath = filename;
            exportDirectory = Path.GetDirectoryName(filePath).Replace(exportDirectoryReplaceFrom, exportDirectoryReplaceTo);
            fileName = Path.GetFileNameWithoutExtension(filePath);

            using ( FileStream stream = File.Open(filename, FileMode.Open, FileAccess.Read) )
            {
                IWorkbook book = ReadXlsFile(stream, filename);

                if ( !IsSupportedFormat(book) ) return true;

                if ( outputBookType == OutputBookType.Individual )
                {
                    ExportScriptableObjectIndividual(book);
                }
                else if ( outputBookType != OutputBookType.Separate )
                {
                    if ( outputBookType == OutputBookType.Dictionary )
                    {
                        if ( !ExportScriptableObjectDictionary(book) )
                        {
                            return false;
                        }
                    }
                    else
                    {
                        if ( !ExportScriptableObjectList(book) )
                        {
                            return false;
                        }
                    }
                }
                else
                {
                    for ( int si = 0 ; si < book.NumberOfSheets ; si++ )
                    {
                        ISheet sheet = book.GetSheetAt(si);
                        if ( !IsValidSheet(sheet) ) continue;    // comment sheet

                        LoadSettings(sheet);

                        if ( GenerateClassParameter(sheet) )
                        {
                            if ( !ExportScriptableObjectSeparate(sheet) )
                            {
                                return false;
                            }
                        }
                    }
                }
            }
            return true;
        }

        // 最初のシートで判定
        bool IsSupportedFormat(IWorkbook book)
        {
            for ( int si = 0 ; si < book.NumberOfSheets ; si++ )
            {
                ISheet sheet = book.GetSheetAt(si);
                if ( !IsValidSheet(sheet) ) continue;    // comment sheet

                IRow row = sheet.GetRow(0);
                if ( row == null ) return false;
                ICell cell = row.GetCell(0);
                if ( cell == null || cell.CellType == CellType.Blank ) return false;
                if ( cell.StringCellValue.Trim() == supportedFormatHeader )
                {
                    LoadSettings(sheet, true);
                    return true;
                }
                break;
            }
            return false;
        }

        void LoadSettings(ISheet sheet, bool isFirst = false)
        {
            startRowNumber = sheet.LastRowNum;
            endRowNumber = sheet.LastRowNum;
            fieldRowNumber = sheet.LastRowNum;
            typeRowNumber = sheet.LastRowNum;
            defaultValueRowNumber = 0;
            subClassName = sheet.SheetName;
            for ( int i = 0 ; i <= sheet.LastRowNum ; i++ )
            {
                IRow row = sheet.GetRow(i);
                if ( row == null ) continue;
                int cellnum = row.PhysicalNumberOfCells;
                if ( cellnum == 0 ) continue;
                ICell cell = row.GetCell(0);
                string cellstr = (cell != null && cell.CellType == CellType.String) ? cell.StringCellValue.Replace("\r", string.Empty).Replace("\n", string.Empty).Trim() : string.Empty;
                if ( cellstr.StartsWith("#") )
                {
                    // Settings
                    string[] settingNameVals = cellstr.Replace("#", string.Empty).Split(';');
                    foreach ( string settingNameVal in settingNameVals )
                    {
                        string[] settingVals = settingNameVal.Split('=');
                        switch ( settingVals[0].Trim().ToLower() )
                        {
                        case "classname":
                            if ( settingVals.Length > 1 )
                            {
                                className = settingVals[1].Trim();
                            }
                            break;
                        case "subclassname":
                            if ( settingVals.Length > 1 )
                            {
                                subClassName = settingVals[1].Trim();
                            }
                            break;
                        case "paramname":
                            if ( settingVals.Length > 1 )
                            {
                                paramStructName = settingVals[1].Trim();
                                if ( paramStructName.Contains(".") )
                                {
                                    isExternalParam = true;
                                }
                                else
                                {
                                    isExternalParam = false;
                                }
                            }
                            break;
                        case "paramfieldname":
                            if ( settingVals.Length > 1 )
                            {
                                paramListFieldName = settingVals[1].Trim();
                            }
                            break;
                        case "sheetstructname":
                            if ( settingVals.Length > 1 )
                            {
                                sheetStructName = settingVals[1].Trim();
                            }
                            break;
                        case "sheetlistfieldname":
                            if ( settingVals.Length > 1 )
                            {
                                sheetListFieldName = settingVals[1].Trim();
                            }
                            break;
                        case "sheetstructnamefieldname":
                            if ( settingVals.Length > 1 )
                            {
                                sheetStructNameFieldName = settingVals[1].Trim();
                            }
                            break;
                        case "sheetstructlistfieldname":
                            if ( settingVals.Length > 1 )
                            {
                                sheetStructListFieldName = settingVals[1].Trim();
                            }
                            break;
                        case "structnameprefix":
                            if ( settingVals.Length > 1 )
                            {
                                structTagName = settingVals[1].Trim();
                            }
                            break;
                        case "subclassnameprefix":
                            if ( settingVals.Length > 1 )
                            {
                                subClassTagName = settingVals[1].Trim();
                            }
                            break;
                        case "subclassfieldnamesuffix":
                            if ( settingVals.Length > 1 )
                            {
                                subClassFieldNameSuffix = settingVals[1].Trim();
                            }
                            break;
                        case "exportpath":
                            if ( settingVals.Length > 1 )
                            {
                                string[] pathnames = settingVals[1].Trim().Split(',');
                                if ( pathnames.Length > 1 )
                                {
                                    exportDirectoryReplaceFrom = pathnames[0].Trim();
                                    exportDirectoryReplaceTo = pathnames[1].Trim();
                                }
                                else
                                {
                                    exportDirectory = settingVals[1].Trim();
                                    exportDirectoryReplaceFrom = string.Empty;
                                    exportDirectoryReplaceTo = string.Empty;
                                }
                            }
                            break;
                        case "separate":
                            if ( isFirst )
                            {
                                outputBookType = OutputBookType.Separate;
                            }
                            break;
                        case "individual":
                            if ( isFirst )
                            {
                                outputBookType = OutputBookType.Individual;
                            }
                            break;
                        case "dictionary":
                            if ( isFirst )
                            {
                                outputBookType = OutputBookType.Dictionary;
                            }
                            break;
                        case "list":
                            if ( isFirst )
                            {
                                outputBookType = OutputBookType.List;
                            }
                            break;
                        case "array":
                            if ( isFirst )
                            {
                                outputBookType = OutputBookType.Array;
                            }
                            break;
                        case "listsheet":
                            if ( isFirst )
                            {
                                outputSheetType = OutputSheetType.List;
                            }
                            break;
                        case "arraysheet":
                            if ( isFirst )
                            {
                                outputSheetType = OutputSheetType.Array;
                            }
                            break;
                        case "noclassdeclaration":
                            if ( !outputClassList.Contains(className) )
                            {
                                outputClassList.Add(className); // 出力済み扱い
                            }
                            break;
                        case "field":
                            if ( cellnum > 1 )
                            {
                                fieldRowNumber = i;
                            }
                            else
                            {
                                fieldRowNumber = ++i;
                            }
                            break;
                        case "type":
                            if ( cellnum > 1 )
                            {
                                typeRowNumber = i;
                            }
                            else
                            {
                                typeRowNumber = ++i;
                            }
                            break;
                        case "default":
                            if ( cellnum > 1 )
                            {
                                defaultValueRowNumber = i;
                            }
                            else
                            {
                                defaultValueRowNumber = ++i;
                            }
                            break;
                        case "start":
                            if ( cellnum > 1 )
                            {
                                startRowNumber = i;
                            }
                            else
                            {
                                startRowNumber = ++i;
                            }
                            break;
                        case "end":
                            if ( cellnum > 1 )
                            {
                                endRowNumber = i;
                            }
                            else
                            {
                                endRowNumber = i - 1;
                            }
                            break;
                        }
                    }
                }
                else if ( startRowNumber > i )
                {
                    if ( fieldRowNumber > i ) fieldRowNumber = i++;
                    if ( typeRowNumber > i ) typeRowNumber = i++;
                    startRowNumber = i;
                }
            }
        }

        bool GenerateClassParameter(ISheet sheet)
        {
            if ( fieldRowNumber == typeRowNumber )
            {
                return false;
            }

            typeList = new List<ExcelColParameter>();
            Dictionary<string, List<string>> fieldNameListList = new Dictionary<string, List<string>>();
            ExcelColParameter lastParser = null;

            IRow fieldRow = sheet.GetRow(fieldRowNumber);
            IRow typeRow = sheet.GetRow(typeRowNumber);
            for ( int i = 0 ; i < fieldRow.LastCellNum ; i++ )
            {
                string fullStructName = string.Empty;
                ICell cell = fieldRow.GetCell(i);
                string fieldName = (cell != null && cell.CellType == CellType.String) ? cell.StringCellValue.Trim() : string.Empty;

                if ( !IsValidFieldName(fieldName) )
                {
                    // Comment row
                }
                else
                {
                    ExcelColParameter parser = new ExcelColParameter();
                    string[] structNames = fieldName.Split('.');
                    parser.structNestCount = structNames.Length - 1;
                    if ( parser.structNestCount > 0 )
                    {
                        parser.structName = structNames[structNames.Length - 2];
                        parser.isStructArray = parser.structName.Contains("[]");
                        if ( parser.isStructArray )
                        {
                            parser.structName = parser.structName.Remove(parser.structName.LastIndexOf("[]"));
                        }
                        fullStructName = fieldName.Remove(fieldName.LastIndexOf("."));
                        fullStructName = fullStructName.Replace("[]", string.Empty);
                    }
                    else
                    {
                        fullStructName = string.Empty;
                        parser.structName = string.Empty;
                    }

                    if ( lastParser != null && lastParser.type == ValueType.UNKNOWN )
                    {
                        if ( lastParser.name == parser.structName )
                        {
                            lastParser.type = ValueType.STRUCT;
                        }
                        else
                        {
                            lastParser.type = ValueType.ENUM;
                        }
                    }

                    if ( fullStructName != string.Empty )
                    {
                        List<string> fieldNameList;
                        if ( fieldNameListList.ContainsKey(fullStructName) )
                        {
                            fieldNameList = fieldNameListList[fullStructName];
                        }
                        else
                        {
                            fieldNameList = new List<string>();
                            fieldNameListList[fullStructName] = fieldNameList;
                        }

                        fieldName = structNames[structNames.Length - 1];
                        if ( fieldNameList.Contains(fieldName) )
                        {
                            parser.isNextStruct = true;
                            fieldNameList.Add(fieldName);
                        }
                        else
                        {
                            fieldNameList.Add(fieldName);
                        }
                    }

                    parser.isArray = fieldName.Contains("[]");
                    if ( parser.isArray )
                    {
                        fieldName = fieldName.Remove(fieldName.LastIndexOf("[]"));
                    }
                    parser.name = fieldName;


                    cell = typeRow.GetCell(i);

                    // array support
                    if ( lastParser != null )
                    {
                        if ( lastParser.isArray && parser.isArray && lastParser.name.Equals(parser.name) )
                        {
                            // trailing array items must be the same as the top type
                            parser.CopyFrom(lastParser);
                            parser.isNextArray = true;
                            parser.colNumber = i;
                            typeList.Add(parser);
                            lastParser = parser;
                            continue;
                        }
                    }

                    if ( cell.CellType != CellType.Unknown && cell.CellType != CellType.Blank )
                    {
                        parser.isEnable = true;

                        try
                        {
                            ValueType vtype = (ValueType)System.Enum.Parse(typeof(ValueType), cell.StringCellValue.ToUpper());
                            parser.type = vtype;
                            parser.typeName = valueTypeNames[(int)vtype];
                            parser.cstypeName = csTypeNames[(int)vtype];
                        }
                        catch
                        {
							Type classType = GetType(cell.StringCellValue);
                            if ( classType != null )
                            {
                                if ( classType.IsEnum )
                                {
                                    parser.type = ValueType.ENUM;
                                }
                                else if ( classType.IsValueType )
                                {
                                    parser.type = ValueType.STRUCT;
                                }
                                else
                                {
                                    parser.type = ValueType.CLASS;
                                }
                            }
                            else
                            {
                                parser.type = ValueType.UNKNOWN;
                            }
                            parser.typeName = cell.StringCellValue;
                            parser.cstypeName = cell.StringCellValue;
                        }
                    }
                    if ( parser.isEnable )
                    {
                        lastParser = parser;
                    }
                    parser.colNumber = i;
                    typeList.Add(parser);
                }
            }
            if ( lastParser != null && lastParser.type == ValueType.UNKNOWN )
            {
                lastParser.type = ValueType.ENUM;
            }
            return true;
        }

        bool ExportClassDefinitionRow(out string outStr, ref List<ExcelColParameter>.Enumerator e, string structName, int structNestCount, string tab)
        {
            StringBuilder builder = new StringBuilder();
            bool hasNext = true;
            string lastStructName = null;

            while ( hasNext )
            {
                ExcelColParameter col = e.Current;
                if ( col.isEnable && !(col.isStructArray && col.isNextStruct) )
                {
                    if ( col.structNestCount < structNestCount || (col.structNestCount == structNestCount && col.structName != structName) )
                    {
                        break;
                    }
                    if ( col.type == ValueType.UNKNOWN )
                    {
                        hasNext = e.MoveNext();
                    }
                    else if ( col.type == ValueType.STRUCT || col.type == ValueType.CLASS )
                    {
                        lastStructName = col.typeName;
                        hasNext = e.MoveNext();
                    }
                    else if ( col.structNestCount > structNestCount )
                    {
                        if ( lastStructName != null && Type.GetType(lastStructName) != null )
                        {
                            ExcelColParameter ncol = e.Current;
                            while ( hasNext && ncol.structNestCount > structNestCount )
                            {
                                hasNext = e.MoveNext();
                                ncol = e.Current;
                            }
                        }
                        else   // define here
                        {
                            if ( lastStructName == null )
                            {
                                lastStructName = structTagName + col.structName;
                            }
                            builder.AppendFormat(tab + "[System.SerializableAttribute]");
                            builder.AppendLine();
                            builder.AppendFormat(tab + "public struct {0}", lastStructName);
                            builder.AppendLine();
                            builder.AppendFormat(tab + "{{");
                            builder.AppendLine();
                            string str;
                            hasNext = ExportClassDefinitionRow(out str, ref e, col.structName, col.structNestCount, tab + "    ");
                            builder.Append(str);
                            builder.AppendFormat(tab + "}}");
                            builder.AppendLine();
                        }
                        if ( col.isStructArray )
                        {
                            builder.AppendFormat(tab + "public {0}[] {1};", lastStructName, col.structName);
                        }
                        else
                        {
                            builder.AppendFormat(tab + "public {0} {1};", lastStructName, col.structName);
                        }
                        builder.AppendLine();
                        lastStructName = null;
                    }
                    else if ( !col.isArray )
                    {
                        builder.AppendFormat(tab + "public {0} {1};", col.typeName, col.name);
                        builder.AppendLine();
                        hasNext = e.MoveNext();
                    }
                    else
                    {
                        if ( !col.isNextArray )
                        {
                            builder.AppendFormat(tab + "public {0}[] {1};", col.typeName, col.name);
                            builder.AppendLine();
                        }
                        hasNext = e.MoveNext();
                    }
                }
                else
                {
                    hasNext = e.MoveNext();
                }
            }

            outStr = builder.ToString();
            return hasNext;
        }

        static string[] classHeaders = new string[] {
                    "using UnityEngine;",
                    "using System.Collections;",
                    "using System.Collections.Generic;",
                    "public class {0} : ScriptableObject",
                    "{",
                };
        static string[] classFooters = new string[] {
                    "}",
                };

        #region Output List
        static string[] fieldHeaders = new string[] {
                    "    public List<{1}> {2};",
                    "",
                };
        static string[] fieldHeaders2 = new string[] {
                    "    [System.SerializableAttribute]",
                    "    public class {1}",
                    "    {",
                };
        static string[] fieldFooters = new string[] {
                    "    }",
                };
        static string[] fieldHeadersArray = new string[] {
                    "    public {3}[] {4};",
                    "",
                    "    [System.SerializableAttribute]",
                    "    public class {3}",
                    "    {",
                    "        public string {5} = string.Empty;",
                    "        public List<{1}> {6};",
                    "    }",
                };
        static string[] fieldHeadersArray2 = new string[] {
                    "    [System.SerializableAttribute]",
                    "    public class {1}",
                    "    {",
                };
        static string[] fieldFootersArray = new string[] {
                    "    }",
                };
        static string[] fieldHeadersList = new string[] {
                    "    public List<{3}> {4};",
                    "",
                    "    [System.SerializableAttribute]",
                    "    public class {3}",
                    "    {",
                    "        public string {5} = string.Empty;",
                    "        public List<{1}> {6};",
                    "    }",
                };
        static string[] fieldHeadersList2 = new string[] {
                    "    [System.SerializableAttribute]",
                    "    public class {1}",
                    "    {",
                };
        static string[] fieldFootersList = new string[] {
                    "    }",
                };
        static string[] fieldHeadersDict = new string[] {
                    "    public Dictionary<string, List<{1}>> {4};",
                    "",
                    "    [System.SerializableAttribute]",
                };
        static string[] fieldHeadersDict2 = new string[] {
                    "    public class {1}",
                    "    {",
                };
        static string[] fieldFootersDict = new string[] {
                    "    }",
                };
        static string[] subClassHeaders = new string[] {
                    "    [System.SerializableAttribute]",
                    "    public class {8}{7}",
                    "    {",
                };
        static string[] subClassFooters = new string[] {
                    "    }",
                    "    public List<{8}{7}> {7}{9};",
                };
        static string[][] outputBookListHeaders = new string[][] { fieldHeaders, fieldHeadersArray, fieldHeadersList, fieldHeadersDict };
        static string[][] outputBookListHeaders2 = new string[][] { fieldHeaders2, fieldHeadersArray2, fieldHeadersList2, fieldHeadersDict2 };
        static string[][] outputBookListFooters = new string[][] { fieldFooters, fieldFootersArray, fieldFootersList, fieldFootersDict };
        #endregion
        #region Output Array
        static string[] arrayFieldHeaders = new string[] {
                    "    public {1}[] {2};",
                    "",
                };
        static string[] arrayFieldHeaders2 = new string[] {
                    "    [System.SerializableAttribute]",
                    "    public class {1}",
                    "    {",
                };
        static string[] arrayFieldFooters = new string[] {
                    "    }",
                };
        static string[] arrayFieldHeadersArray = new string[] {
                    "    public {3}[] {4};",
                    "",
                    "    [System.SerializableAttribute]",
                    "    public class {3}",
                    "    {",
                    "        public string {5} = string.Empty;",
                    "        public {1}[] {6};",
                    "    }",
                };
        static string[] arrayFieldHeadersArray2 = new string[] {
                    "    [System.SerializableAttribute]",
                    "    public class {1}",
                    "    {",
                };
        static string[] arrayFieldFootersArray = new string[] {
                    "    }",
                };
        static string[] arrayFieldHeadersList = new string[] {
                    "    public List<{3}> {4};",
                    "",
                    "    [System.SerializableAttribute]",
                    "    public class {3}",
                    "    {",
                    "        public string {5} = string.Empty;",
                    "        public {1}[] {6};",
                    "    }",
                };
        static string[] arrayFieldHeadersList2 = new string[] {
                    "    [System.SerializableAttribute]",
                    "    public class {1}",
                    "    {",
                };
        static string[] arrayFieldFootersList = new string[] {
                    "    }",
                };
        static string[] arrayFieldHeadersDict = new string[] {
                    "    public Dictionary<string, {1}[]> {4};",
                    "",
                    "    [System.SerializableAttribute]",
                };
        static string[] arrayFieldHeadersDict2 = new string[] {
                    "    public class {1}",
                    "    {",
                };
        static string[] arrayFieldFootersDict = new string[] {
                    "    }",
                };
        static string[] arraySubClassHeaders = new string[] {
                    "    [System.SerializableAttribute]",
                    "    public class {8}{7}",
                    "    {",
                };
        static string[] arraySubClassFooters = new string[] {
                    "    }",
                    "    public {8}{7}[] {7}{9};",
                };
        static string[][] outputBookArrayHeaders = new string[][] { arrayFieldHeaders, arrayFieldHeadersArray, arrayFieldHeadersList, arrayFieldHeadersDict };
        static string[][] outputBookArrayHeaders2 = new string[][] { arrayFieldHeaders2, arrayFieldHeadersArray2, arrayFieldHeadersList2, arrayFieldHeadersDict2 };
        static string[][] outputBookArrayFooters = new string[][] { arrayFieldFooters, arrayFieldFootersArray, arrayFieldFootersList, arrayFieldFootersDict };
        #endregion

        bool ExportClassDefinition(ISheet sheet)
        {
            if ( outputClassList.Contains(className) )  // already output
            {
                return false;
            }
            string[] fheaders;
            string[] fheaders2;
            string[] ffooters;
            if ( outputSheetType == OutputSheetType.Array )
            {
                fheaders = outputBookArrayHeaders[(int)outputBookType];
                fheaders2 = outputBookArrayHeaders2[(int)outputBookType];
                ffooters = outputBookArrayFooters[(int)outputBookType];
            }
            else
            {
                fheaders = outputBookListHeaders[(int)outputBookType];
                fheaders2 = outputBookListHeaders2[(int)outputBookType];
                ffooters = outputBookListFooters[(int)outputBookType];
            }
            StringBuilder builder = new StringBuilder();

            string[] replaceStrings = new string[] { className, paramStructName, paramListFieldName, sheetStructName, sheetListFieldName, sheetStructNameFieldName, sheetStructListFieldName };

            foreach ( string str in classHeaders )
            {
                string rstr = str;
                for ( int i = 0 ; i < replaceStrings.Length ; i++ )
                {
                    rstr = rstr.Replace("{" + i.ToString() + "}", replaceStrings[i]);
                }
                builder.Append(rstr);
                builder.AppendLine();
            }
            foreach ( string str in fheaders )
            {
                string rstr = str;
                for ( int i = 0 ; i < replaceStrings.Length ; i++ )
                {
                    rstr = rstr.Replace("{" + i.ToString() + "}", replaceStrings[i]);
                }
                builder.Append(rstr);
                builder.AppendLine();
            }

            if ( !isExternalParam )
            {
                foreach ( string str in fheaders2 )
                {
                    string rstr = str;
                    for ( int i = 0 ; i < replaceStrings.Length ; i++ )
                    {
                        rstr = rstr.Replace("{" + i.ToString() + "}", replaceStrings[i]);
                    }
                    builder.Append(rstr);
                    builder.AppendLine();
                }

                List<ExcelColParameter>.Enumerator e = typeList.GetEnumerator();
                e.MoveNext();
                string exportStr;
                ExportClassDefinitionRow(out exportStr, ref e, string.Empty, 0, "        ");
                builder.Append(exportStr);

                foreach ( string str in ffooters )
                {
                    string rstr = str;
                    for ( int i = 0 ; i < replaceStrings.Length ; i++ )
                    {
                        rstr = rstr.Replace("{" + i.ToString() + "}", replaceStrings[i]);
                    }
                    builder.Append(rstr);
                    builder.AppendLine();
                }
            }

            foreach ( string str in classFooters )
            {
                string rstr = str;
                for ( int i = 0 ; i < replaceStrings.Length ; i++ )
                {
                    rstr = rstr.Replace("{" + i.ToString() + "}", replaceStrings[i]);
                }
                builder.Append(rstr);
                builder.AppendLine();
            }

            string filename = classExportPath + className + ".cs";
            if ( ExportTextFile(classExportPath, filename, builder.ToString()) )
            {
                ExcelScriptableObjectImporter.AddGenerateClassFileName(filename);
            }
            outputClassList.Add(className);
            return true;
        }

        bool ExportClassDefinitionIndividual(IWorkbook book)
        {
            if ( outputClassList.Contains(className) )  // already output
            {
                return false;
            }
            string[] cheaders;
            string[] cfooters;
            if ( outputSheetType == OutputSheetType.Array )
            {
                cheaders = arraySubClassHeaders;
                cfooters = arraySubClassFooters;
            }
            else
            {
                cheaders = subClassHeaders;
                cfooters = subClassFooters;
            }
            StringBuilder builder = new StringBuilder();

			string[] replaceStrings = new string[] { className, paramStructName, paramListFieldName, sheetStructName, sheetListFieldName, sheetStructNameFieldName, sheetStructListFieldName, subClassName, subClassTagName, subClassFieldNameSuffix};

            foreach ( string str in classHeaders )
            {
                string rstr = str;
                for ( int i = 0 ; i < replaceStrings.Length ; i++ )
                {
                    rstr = rstr.Replace("{" + i.ToString() + "}", replaceStrings[i]);
                }
                builder.Append(rstr);
                builder.AppendLine();
            }

            for ( int si = 0 ; si < book.NumberOfSheets ; si++ )
            {
                ISheet sheet = book.GetSheetAt(si);
                if ( !IsValidSheet(sheet) ) continue;    // comment sheet

                LoadSettings(sheet);
                GenerateClassParameter(sheet);

				replaceStrings = new string[] { className, paramStructName, paramListFieldName, sheetStructName, sheetListFieldName, sheetStructNameFieldName, sheetStructListFieldName, subClassName, subClassTagName, subClassFieldNameSuffix};

                foreach ( string str in cheaders )
                {
	                string rstr = str;
	                for ( int i = 0 ; i < replaceStrings.Length ; i++ )
	                {
	                    rstr = rstr.Replace("{" + i.ToString() + "}", replaceStrings[i]);
	                }
	                builder.Append(rstr);
	                builder.AppendLine();
                }
                
                List<ExcelColParameter>.Enumerator e = typeList.GetEnumerator();
                e.MoveNext();
                string exportStr = string.Empty;
                ExportClassDefinitionRow(out exportStr, ref e, string.Empty, 0, "           ");
                builder.Append(exportStr);

                foreach ( string str in cfooters )
                {
	                string rstr = str;
	                for ( int i = 0 ; i < replaceStrings.Length ; i++ )
	                {
	                    rstr = rstr.Replace("{" + i.ToString() + "}", replaceStrings[i]);
	                }
	                builder.Append(rstr);
	                builder.AppendLine();
                }

            }

            foreach ( string str in classFooters )
            {
                string rstr = str;
                for ( int i = 0 ; i < replaceStrings.Length ; i++ )
                {
                    rstr = rstr.Replace("{" + i.ToString() + "}", replaceStrings[i]);
                }
                builder.Append(rstr);
                builder.AppendLine();
            }

            string filename = classExportPath + className + ".cs";
            if ( ExportTextFile(classExportPath, filename, builder.ToString()) )
            {
                ExcelScriptableObjectImporter.AddGenerateClassFileName(filename);
            }
            outputClassList.Add(className);
            return true;
        }

        public Type GetType(string typeName)
        {
            Type type = Type.GetType(typeName);

            if ( type != null )
                return type;

            string unitytypeName = typeName + ",Assembly-CSharp";
            type = Type.GetType(unitytypeName);
            if (type != null)
                return type;

            string cstypeName = typeName.Replace(".", "+");

            var referencedAssemblies = AppDomain.CurrentDomain.GetAssemblies();
            foreach ( var assembly in referencedAssemblies )
            {
                if ( assembly != null )
                {
                    type = assembly.GetType(cstypeName);
                    if ( type != null )
                        return type;
                }
            }

            unitytypeName = typeName + ",UnityEngine";
            type = Type.GetType(unitytypeName);
            if ( type != null )
                return type;

            unitytypeName = "UnityEngine." + typeName + ",UnityEngine";
            type = Type.GetType(unitytypeName);
            if (type != null)
                return type;

            return null;
        }

        float[] GetFloatArrayValue(string str, int count, float defaultValue)
        {
            string[] sval = str.Split(',');
            float[] array = new float[count];
            int len = Mathf.Min(sval.Length, count);
            int i = 0;
            for ( ; i < len ; i++ )
            {
                try
                {
                    array[i] = float.Parse(sval[i]);
                }
                catch
                {
                    array[i] = defaultValue;
                }
            }
            for ( ; i < len ; i++ )
            {
                array[i] = defaultValue;
            }
            return array;
        }

        object GetCellValue(IRow row, IRow defrow, ExcelColParameter col)
        {
            ICell defcell = defrow != null ? defrow.GetCell(col.colNumber) : null;
            ICell cell = row.GetCell(col.colNumber);
            switch ( col.type )
            {
            case ValueType.BOOL:
                return (cell == null ? (defcell == null ? false : defcell.BooleanCellValue) : cell.BooleanCellValue);
            case ValueType.DOUBLE:
                return (double)(cell == null ? (defcell == null ? 0 : defcell.NumericCellValue) : cell.NumericCellValue);
            case ValueType.INT:
                return (int)(cell == null ? (defcell == null ? 0 : defcell.NumericCellValue) : cell.NumericCellValue);
            case ValueType.FLOAT:
                return (float)(cell == null ? (defcell == null ? 0 : defcell.NumericCellValue) : cell.NumericCellValue);
            case ValueType.STRING:
                return (string)(cell == null ? (defcell == null ? string.Empty : defcell.StringCellValue) : cell.StringCellValue);
            case ValueType.VECTOR2:
                {
                    float[] fval = GetFloatArrayValue((string)(cell == null ? (defcell == null ? string.Empty : defcell.StringCellValue) : cell.StringCellValue), 2, 0.0f);
                    Vector2 v2 = new Vector2(fval[0], fval[1]);
                    return v2;
                }
            case ValueType.VECTOR3:
                {
                    float[] fval = GetFloatArrayValue((string)(cell == null ? (defcell == null ? string.Empty : defcell.StringCellValue) : cell.StringCellValue), 3, 0.0f);
                    Vector3 v3 = new Vector3(fval[0], fval[1], fval[2]);
                    return v3;
                }
            case ValueType.COLOR:
                {
                    float[] fval = GetFloatArrayValue((string)(cell == null ? (defcell == null ? string.Empty : defcell.StringCellValue) : cell.StringCellValue), 4, 1.0f);
                    Color c = new Color(fval[0], fval[1], fval[2], fval[3]);
                    return c;
                }
            case ValueType.HASH:
				{
					string str = (string)(cell == null ? (defcell == null ? string.Empty : defcell.StringCellValue) : cell.StringCellValue);
					int id = 0;
					if (str != null && str != string.Empty && str != "")
					{
						// アニメータのハッシュにすると、そのままアニメ再生に流用できる。
						id = Animator.StringToHash(str);
						// id = NetworkInstance.GetHashCode(str);
						if (id == 0) Debug.LogError("[" + str + "] is zero!");
					}
	                return id;
				}

			case ValueType.ENUM:
                Type enumType = GetType(col.cstypeName);
				object o = null;
                try
                {
//					o = (cell == null ? (defcell == null ? Convert.ChangeType(0, enumType) : System.Enum.Parse(enumType, defcell.StringCellValue)) : System.Enum.Parse(enumType, cell.StringCellValue));
                    string str = cell == null ? (defcell == null ? "" : defcell.StringCellValue) : cell.StringCellValue;
                    o = (str == "" ? System.Enum.ToObject(enumType, 0) : System.Enum.Parse(enumType, str));
                }
                catch
                {
					if (cell == null)
					{
						if (defcell == null)
						{
							Debug.LogError("Convert Error: 0 to " + col.cstypeName);
						}
						else
						{
							Debug.LogError("Convert Error: " + defcell.StringCellValue + " to " + col.cstypeName);
						}
					}
					else
					{
						Debug.LogError("Convert Error: " + cell.StringCellValue + " to " + col.cstypeName);
					}
                }
                return o;
            }
            return null;
        }

        bool ExportScriptableObjectRow(object obj, ref List<ExcelColParameter>.Enumerator e, IRow row, IRow lastrow, IRow defrow, string structName, string fullStructName, string fullStructTagName, int structNestCount, int arrayNestCount, out int valueCount)
        {
            List<string> nameList = new List<string>();
            bool hasNext = true;
            string lastStructName = null;
            valueCount = 0;

            while ( hasNext )
            {
                ExcelColParameter col = e.Current;
                if ( col.isEnable )
                {
                    if ( col.structNestCount < structNestCount || (col.structNestCount == structNestCount && (col.structName != structName || nameList.Contains(col.name))) )
                    {
                        break;
                    }
                    if ( col.type == ValueType.UNKNOWN )
                    {
                        hasNext = e.MoveNext();
                    }
                    else if ( col.type == ValueType.STRUCT || col.type == ValueType.CLASS )
                    {
                        lastStructName = col.typeName;
                        hasNext = e.MoveNext();
                    }
                    else if ( col.structNestCount > structNestCount || (col.isStructArray && col.structNestCount > arrayNestCount) )
                    {
                        if ( lastStructName != null )
                        {
                            if ( Type.GetType(lastStructName) == null )
                            {
                                lastStructName = fullStructTagName + (fullStructName == string.Empty ? string.Empty : ".") + (col.structName == string.Empty ? string.Empty : lastStructName);
                            }
                        }
                        else
                        {
                            lastStructName = fullStructTagName + (fullStructName == string.Empty ? string.Empty : ".") + (col.structName == string.Empty ? string.Empty : structTagName + col.structName);
                        }
                        string newFullStructName = fullStructName + (fullStructName == string.Empty ? string.Empty : ".") + col.structName;
                        string newFullStructTagName = lastStructName;
                        if ( col.isStructArray )
                        {
                            Type structType = GetType(newFullStructTagName);
                            IList fieldList = (IList)Activator.CreateInstance(typeof(List<>).MakeGenericType(structType));
                            List<ExcelColParameter>.Enumerator ea = e;
                            do
                            {
                                object newObj = Activator.CreateInstance(structType);
                                IRow lrow = GetLastMultilineRow(row, col);
                                int valCount;
                                hasNext = ExportScriptableObjectRow(newObj, ref ea, row, lrow, defrow, col.structName, newFullStructName, newFullStructTagName, col.structNestCount, col.isStructArray ? col.structNestCount : 0, out valCount);
                                if ( !isIgnoredEmptyElement || valCount > 0 )
                                {
                                    fieldList.Add(newObj);
                                }
							}
                            while ( hasNext && col.structNestCount == ea.Current.structNestCount && col.structName == ea.Current.structName );

                            IRow nextrow = GetNextValueRow(row, col);
                            while ( nextrow != null && nextrow.RowNum <= lastrow.RowNum )
                            {
                                ea = e;
                                do
                                {
                                    object newObj = Activator.CreateInstance(structType);
                                    IRow lrow = GetLastMultilineRow(nextrow, col);
                                    int valCount;
                                    hasNext = ExportScriptableObjectRow(newObj, ref ea, nextrow, lrow, defrow, col.structName, newFullStructName, newFullStructTagName, col.structNestCount, col.isStructArray ? col.structNestCount : 0, out valCount);
                                    if ( !isIgnoredEmptyElement || valCount > 0 )
                                    {
                                        fieldList.Add(newObj);
                                    }
                                }
                                while ( hasNext && col.structNestCount == ea.Current.structNestCount && col.structName == ea.Current.structName );
                                nextrow = GetNextValueRow(nextrow, col);
                            }

                            e = ea;

                            Array array = Array.CreateInstance(structType, fieldList.Count);
                            fieldList.CopyTo(array, 0);
                            obj.GetType().GetField(col.structName).SetValue(obj, array);
                        }
                        else
                        {
                            Type structType = GetType(newFullStructTagName);
                            object newObj = Activator.CreateInstance(structType);
                            IRow lrow = GetLastMultilineRow(row, col);
                            int valCount;
                            hasNext = ExportScriptableObjectRow(newObj, ref e, row, lrow, defrow, col.structName, newFullStructName, newFullStructTagName, col.structNestCount, col.isStructArray ? col.structNestCount : 0, out valCount);
                            obj.GetType().GetField(col.structName).SetValue(obj, newObj);
                        }
                        lastStructName = null;
                    }
                    else
                    {
                        if ( !col.isArray )
                        {
                            if ( IsValidParameterCell(row.GetCell(col.colNumber)) ) valueCount++;
                            obj.GetType().GetField(col.name).SetValue(obj, GetCellValue(row, defrow, col));

                            hasNext = e.MoveNext();
                            nameList.Add(col.name);
                        }
                        else
                        {
                            Type fieldType = GetType(col.cstypeName);
                            if ( fieldType == null ) Debug.LogWarning("Type Error: " + col.cstypeName);
                            IList fieldList = (IList)Activator.CreateInstance(typeof(List<>).MakeGenericType(fieldType));

                            string colname = col.name;
                            nameList.Add(colname);

                            List<ExcelColParameter>.Enumerator ea = e;
                            ExcelColParameter colarray = ea.Current;
                            do
                            {
                                bool bvalid = IsValidParameterCell(row.GetCell(col.colNumber));
                                if ( bvalid ) valueCount++;
                                if ( !isIgnoredEmptyElement || bvalid )
                                {
                                    fieldList.Add(GetCellValue(row, defrow, colarray));
                                }

                                hasNext = ea.MoveNext();
                                colarray = ea.Current;
                                if ( !colarray.isNextArray )
                                {
                                    break;
                                }
                            }
                            while ( hasNext );

                            IRow nextrow = GetNextMultilineRow(row, lastrow, col);
                            while ( nextrow != null && nextrow.RowNum <= lastrow.RowNum )
                            {
                                ea = e;
                                colarray = ea.Current;
                                do
                                {
                                    bool bvalid = IsValidParameterCell(row.GetCell(col.colNumber));
                                    if ( bvalid ) valueCount++;
                                    if ( !isIgnoredEmptyElement || bvalid )
                                    {
                                        fieldList.Add(GetCellValue(nextrow, defrow, colarray));
                                    }

                                    hasNext = ea.MoveNext();
                                    colarray = ea.Current;
                                    if ( !colarray.isNextArray )
                                    {
                                        break;
                                    }
                                }
                                while ( hasNext );
                                nextrow = GetNextMultilineRow(nextrow, lastrow, col);
                            }

                            e = ea;

                            Array array = Array.CreateInstance(fieldType, fieldList.Count);
                            fieldList.CopyTo(array, 0);
                            obj.GetType().GetField(colname).SetValue(obj, array);
                        }
                    }
                }
                else
                {
                    hasNext = e.MoveNext();
                }
            }
            if ( valueCount == 0 )
            {

            }
            return hasNext;
        }

        IRow GetPreviousRow(IRow row)
        {
            int rownum = row.RowNum;
            if ( rownum == 0 || row.Sheet.LastRowNum < rownum )
            {
                return row.Sheet.GetRow(row.Sheet.LastRowNum);
            }
            return row.Sheet.GetRow(rownum-1);
        }

        IRow GetNextValueRow(IRow row, ExcelColParameter col)
        {
            int rownum = row.RowNum;
            while ( row.Sheet.LastRowNum > rownum )
            {
                IRow nextrow = row.Sheet.GetRow(++rownum);
                if ( !IsValidRow(nextrow) ) continue;
                ICell cell = nextrow.GetCell(col.colNumber);
                if ( IsValidParameterCell(cell) )
                {
                    return nextrow;
                }
            }
            return null;
        }

        IRow GetNextBlankRow(IRow row, ExcelColParameter col)
        {
            int rownum = row.RowNum;
            while ( row.Sheet.LastRowNum > rownum )
            {
                IRow nextrow = row.Sheet.GetRow(++rownum);
                if ( !IsCommentRow(nextrow) ) continue;
                ICell cell = nextrow.GetCell(col.colNumber);
                if ( !IsValidParameterCell(cell) )
                {
                    return nextrow;
                }
            }
            return null;
        }

        IRow GetLastMultilineRow(IRow row, ExcelColParameter col)
        {
            if ( col.isArray )
            {
                IRow lastrow = GetNextBlankRow(row, col);
                if ( lastrow == null ) return row.Sheet.GetRow(row.Sheet.LastRowNum);
                lastrow = GetNextValueRow(lastrow, col);
                return (lastrow == null) ? row.Sheet.GetRow(row.Sheet.LastRowNum) : GetPreviousRow(lastrow);
            }
            else
            {
                IRow lastrow = GetNextValueRow(row, col);
                return (lastrow == null) ? row.Sheet.GetRow(row.Sheet.LastRowNum) : GetPreviousRow(lastrow);
            }
        }

        IRow GetNextMultilineRow(IRow row, IRow lastrow, ExcelColParameter col)
        {
            int rownum = row.RowNum;
            int lastrownum = (lastrow == null) ? row.Sheet.LastRowNum : lastrow.RowNum;
            while ( lastrownum > rownum )
            {
                IRow nextrow = row.Sheet.GetRow(++rownum);
                if ( !IsValidRow(nextrow) ) continue;
                ICell cell = nextrow.GetCell(col.colNumber);
                if ( !IsValidParameterCell(cell) )
                {
                    break;
                }
                return nextrow;
            }
            return null;
        }

        ExcelColParameter GetPrevColParameter(ExcelColParameter col)
        {
            int index = typeList.FindIndex(x => x == col);
            if ( index > 1 )
            {
                return typeList[index - 1];
            }
            return null;
        }

        IList CreateParamList(ISheet sheet, string structname, Type paramType)
        {
            Type paramListType = typeof(List<>).MakeGenericType(paramType);
            IList paramList = (IList)Activator.CreateInstance(paramListType);

            IRow defaultrow = defaultValueRowNumber > 0 ? sheet.GetRow(defaultValueRowNumber) : null;
            for ( int i = startRowNumber ; i <= endRowNumber ; i++ )
            {
                IRow row = sheet.GetRow(i);

                if ( !IsValidRow(row) ) continue;    // comment

                if ( !IsValidParameterRow(row) ) continue;

                object p = Activator.CreateInstance(paramType);

                List<ExcelColParameter>.Enumerator e = typeList.GetEnumerator();
                e.MoveNext();
                IRow lastrow = GetLastMultilineRow(row, e.Current);
                int valCount;
                ExportScriptableObjectRow(p, ref e, row, lastrow, defaultrow, string.Empty, className + "." + structname, className + "." + structname, 0, 0, out valCount);

                i = lastrow.RowNum;

                paramList.Add(p);
            }

            return paramList;
        }

        bool IsValidFieldName(string name)
        {
            if ( name == string.Empty || name.StartsWith("#") )
            {
                return false;   // Comment
            }
            return true;
        }

        bool IsValidSheet(ISheet sheet)
        {
            if ( sheet.SheetName.Trim() == string.Empty || sheet.SheetName.StartsWith("#") )
            {
                return false;    // Comment
            }
            return true;
        }

        bool IsCommentCell(ICell cell)
        {
            if ( cell == null || cell.CellType != CellType.String) return false;

            string str = cell.StringCellValue.Trim();
            if ( str.Equals("#") || str.StartsWith("# "))
            {
                return true;    // Comment
            }
            return false;
        }

        bool IsCommentRow(IRow row)
        {
            if ( row == null ) return false;

            ICell cell = row.GetCell(0);
            if ( IsCommentCell(cell) ) return false;    // comment

            return true;
        }

        bool IsValidRow(IRow row)
        {
            if ( row == null || row.PhysicalNumberOfCells == 0 ) return false;

            ICell cell = row.GetCell(0);
            if ( IsCommentCell(cell) ) return false;    // comment

            return true;
        }

        bool IsValidParameterCell(ICell cell)
        {
            if ( cell != null && cell.CellType != CellType.Blank )
            {
                if ( cell.CellType != CellType.String || (cell.CellType == CellType.String && !cell.StringCellValue.Trim().Equals(string.Empty)) )
                {
                    return true;
                }
            }
            return false;
        }

        bool IsValidParameterRow(IRow row)
        {
            foreach ( ExcelColParameter cp in typeList )
            {
                ICell cell = row.GetCell(cp.colNumber);
                if ( IsValidParameterCell(cell) ) return true;
            }
            return false;
        }

        Type GetParameterType(string structname)
        {
            return GetType(structname.Contains(".") ? (structname) : (className + "." + structname));
        }

        bool ExportScriptableObjectSeparate(ISheet sheet)
        {
            Type classType = GetType(className);
            if ( classType == null ) return false;

            string exportFileName = sheet.SheetName;
            string exportPath = exportDirectory + "/" + exportFileName + ".asset";

            object data = AssetDatabase.LoadAssetAtPath(exportPath, classType);
            if ( data == null )
            {
                data = ScriptableObject.CreateInstance(classType);
                Directory.CreateDirectory(exportDirectory);
                AssetDatabase.CreateAsset((ScriptableObject)data, exportPath);
            }

            Type paramType = GetParameterType(paramStructName);
            IList paramList = CreateParamList(sheet, paramStructName, paramType);

            if ( outputSheetType == OutputSheetType.Array )
            {
                Array array = Array.CreateInstance(paramType, paramList.Count);
                paramList.CopyTo(array, 0);
                data.GetType().GetField(paramListFieldName).SetValue(data, array);
            }
            else
            {
                data.GetType().GetField(paramListFieldName).SetValue(data, paramList);
            }
            EditorUtility.SetDirty((ScriptableObject)data);
            return true;
        }

        bool ExportScriptableObjectDictionary(IWorkbook book)
        {
            Type classType = GetType(className);
            if ( classType == null ) return false;

            string exportFileName = fileName;
            string exportPath = exportDirectory + "/" + exportFileName + ".asset";

            object data = AssetDatabase.LoadAssetAtPath(exportPath, classType);
            if ( data == null )
            {
                data = ScriptableObject.CreateInstance(classType);
                Directory.CreateDirectory(exportDirectory);
                AssetDatabase.CreateAsset((ScriptableObject)data, exportPath);
            }

            object paramDict;
            Type paramType = GetParameterType(paramStructName);
            if ( outputSheetType == OutputSheetType.Array )
            {
                Type paramArrayType = paramType.MakeArrayType();
                Type paramDictType = typeof(Dictionary<,>).MakeGenericType(typeof(string), paramArrayType);
                paramDict = Activator.CreateInstance(paramDictType);

                for ( int si = 0 ; si < book.NumberOfSheets ; si++ )
                {
                    ISheet sheet = book.GetSheetAt(si);
                    if ( !IsValidSheet(sheet) ) continue;    // comment sheet

                    LoadSettings(sheet);
                    IList paramList = CreateParamList(sheet, paramStructName, paramType);

                    Array array = Array.CreateInstance(paramType, paramList.Count);
                    paramList.CopyTo(array, 0);
                    paramDict.GetType().GetMethod("Add").Invoke(paramDict, new object[] { sheet.SheetName.Trim(), array });
                }
            }
            else
            {
                Type paramListType = typeof(List<>).MakeGenericType(paramType);
                Type paramDictType = typeof(Dictionary<,>).MakeGenericType(typeof(string), paramListType);
                paramDict = Activator.CreateInstance(paramDictType);

                for ( int si = 0 ; si < book.NumberOfSheets ; si++ )
                {
                    ISheet sheet = book.GetSheetAt(si);
                    if ( !IsValidSheet(sheet) ) continue;    // comment sheet

                    LoadSettings(sheet);
                    IList paramList = CreateParamList(sheet, paramStructName, paramType);

                    paramDict.GetType().GetMethod("Add").Invoke(paramDict, new object[] { sheet.SheetName.Trim(), paramList });
                }
            }

            data.GetType().GetField(sheetListFieldName).SetValue(data, paramDict);
            EditorUtility.SetDirty((ScriptableObject)data);
            return true;
        }

        bool ExportScriptableObjectList(IWorkbook book)
        {
            Type classType = GetType(className);
            if ( classType == null ) return false;

            string exportFileName = fileName;
            string exportPath = exportDirectory + "/" + exportFileName + ".asset";

            object data = AssetDatabase.LoadAssetAtPath(exportPath, classType);
            if ( data == null )
            {
                data = ScriptableObject.CreateInstance(classType);
                Directory.CreateDirectory(exportDirectory);
                AssetDatabase.CreateAsset((ScriptableObject)data, exportPath);
            }

            Type sheetType = GetType(className + "." + sheetStructName);
            Type sheetListType = typeof(List<>).MakeGenericType(sheetType);
            IList sheetList = (IList)Activator.CreateInstance(sheetListType);
            Type paramType = GetParameterType(paramStructName);

            for ( int si = 0 ; si < book.NumberOfSheets ; si++ )
            {
                ISheet sheet = book.GetSheetAt(si);
                if ( !IsValidSheet(sheet) ) continue;    // comment sheet

                LoadSettings(sheet);
                IList paramList = CreateParamList(sheet, paramStructName, paramType);

                object sheetdata = Activator.CreateInstance(sheetType);
                sheetdata.GetType().GetField(sheetStructNameFieldName).SetValue(sheetdata, sheet.SheetName.Trim());
                if ( outputSheetType == OutputSheetType.Array )
                {
                    Array array = Array.CreateInstance(paramType, paramList.Count);
                    paramList.CopyTo(array, 0);
                    sheetdata.GetType().GetField(sheetStructListFieldName).SetValue(sheetdata, array);
                }
                else
                {
                    sheetdata.GetType().GetField(sheetStructListFieldName).SetValue(sheetdata, paramList);
                }
                sheetList.Add(sheetdata);
            }

            if ( outputBookType == OutputBookType.Array )
            {
                Array array = Array.CreateInstance(sheetType, sheetList.Count);
                sheetList.CopyTo(array, 0);
                data.GetType().GetField(sheetListFieldName).SetValue(data, array);
            }
            else
            {
                data.GetType().GetField(sheetListFieldName).SetValue(data, sheetList);
            }
            EditorUtility.SetDirty((ScriptableObject)data);
            return true;
        }

        bool ExportScriptableObjectIndividual(IWorkbook book)
        {
            Type classType = GetType(className);
            if ( classType == null ) return false;

            string exportFileName = fileName;
            string exportPath = exportDirectory + "/" + exportFileName + ".asset";

            object data = AssetDatabase.LoadAssetAtPath(exportPath, classType);
            if ( data == null )
            {
                data = ScriptableObject.CreateInstance(classType);
                Directory.CreateDirectory(exportDirectory);
                AssetDatabase.CreateAsset((ScriptableObject)data, exportPath);
            }

            for ( int si = 0 ; si < book.NumberOfSheets ; si++ )
            {
                ISheet sheet = book.GetSheetAt(si);
                if ( !IsValidSheet(sheet) ) continue;    // comment sheet

                LoadSettings(sheet);
                GenerateClassParameter(sheet);
                Type paramType = GetParameterType(subClassTagName + subClassName);
                IList paramList = CreateParamList(sheet, subClassTagName + subClassName, paramType);

                if ( outputSheetType == OutputSheetType.Array )
                {
                    Array array = Array.CreateInstance(paramType, paramList.Count);
                    paramList.CopyTo(array, 0);
                    data.GetType().GetField(subClassName + subClassFieldNameSuffix).SetValue(data, array);
                }
                else
                {
                    data.GetType().GetField(subClassName + subClassFieldNameSuffix).SetValue(data, paramList);
                }
            }

            EditorUtility.SetDirty((ScriptableObject)data);
            return true;
        }


        bool ExportTextFile(string path, string fullpathname, string outStr)
        {
            if ( File.Exists(fullpathname) )
            {
                string inStr = File.ReadAllText(fullpathname);
                if ( inStr.Equals(outStr) )
                {
                    return false;
                }
                FileAttributes attr = File.GetAttributes(fullpathname);
                if ( (attr & FileAttributes.ReadOnly) != 0 )
                {
                    // ToDo: checkout
                    attr = attr & ~FileAttributes.ReadOnly;
                    File.SetAttributes(fullpathname, attr);
                }
            }
            else
            {
                Directory.CreateDirectory(path);
            }
            File.WriteAllText(fullpathname, outStr);
            return true;
        }

    }
}
