using System;
using System.IO;
using System.Data;
using System.Text;
using System.Collections.Generic;
using Newtonsoft.Json;

namespace excel2json
{
    /// <summary>
    /// 将DataTable对象，转换成JSON string，并保存到文件中
    /// </summary>
    class JsonExporter
    {
        string mContext = "";
        int mHeaderRows = 0;
        Dictionary<string, object> mData = new Dictionary<string, object>();
        JsonSerializerSettings jsonSettings;

        public string context {
            get {
                return mContext;
            }
        }

        /// <summary>
        /// 构造函数：完成内部数据创建
        /// </summary>
        /// <param name="excel">ExcelLoader Object</param>
        public JsonExporter(ExcelLoader excel, bool lowcase, bool exportArray, string dateFormat, bool forceSheetName, int headerRows, string excludePrefix, bool cellJson)
        {
            mHeaderRows = 4 - 1;
            List<DataTable> validSheets = new List<DataTable>();
            for (int i = 0; i < excel.Sheets.Count; i++)
            {
                DataTable sheet = excel.Sheets[i];

                // 过滤掉包含特定前缀的表单
                string sheetName = sheet.TableName;
                if (excludePrefix.Length > 0 && sheetName.StartsWith(excludePrefix))
                    continue;

                if (sheet.Columns.Count > 0 && sheet.Rows.Count > 0 && sheet.Rows[1][0].ToString() == "tid")
                    validSheets.Add(sheet);
            }

            jsonSettings = new JsonSerializerSettings
            {
                DateFormatString = dateFormat,
                Formatting = Formatting.Indented
            };

            if (!forceSheetName && validSheets.Count == 1 && false)
            {   // single sheet

                //-- convert to object
                object sheetValue = convertSheet(validSheets[0], exportArray, lowcase, excludePrefix, cellJson);

                //-- convert to json string
                mContext = JsonConvert.SerializeObject(sheetValue, jsonSettings);
            }
            else
            { // mutiple sheet

                foreach (var sheet in validSheets)
                {
                    Dictionary<string, object> data = new Dictionary<string, object>();
                    object sheetValue = convertSheet(sheet, exportArray, lowcase, excludePrefix, cellJson);
                    data.Add("data", sheetValue);
                    mData.Add(sheet.TableName, data);
                }

                //-- convert to json string
                mContext = JsonConvert.SerializeObject(mData, jsonSettings);
            }
        }

        private object convertSheet(DataTable sheet, bool exportArray, bool lowcase, string excludePrefix, bool cellJson)
        {
            if (exportArray)
                return convertSheetToArray(sheet, lowcase, excludePrefix, cellJson);
            else
                return convertSheetToDict(sheet, lowcase, excludePrefix, cellJson);
        }

        private object convertSheetToArray(DataTable sheet, bool lowcase, string excludePrefix, bool cellJson)
        {
            List<object> values = new List<object>();

            int firstDataRow = mHeaderRows;
            for (int i = firstDataRow; i < sheet.Rows.Count; i++)
            {
                DataRow row = sheet.Rows[i];

                values.Add(
                    convertRowToDict(sheet, row, lowcase, firstDataRow, excludePrefix, cellJson)
                    );
            }

            return values;
        }

        /// <summary>
        /// 以第一列为ID，转换成ID->Object的字典对象
        /// </summary>
        private object convertSheetToDict(DataTable sheet, bool lowcase, string excludePrefix, bool cellJson)
        {
            Dictionary<int, object> importData =
                new Dictionary<int, object>();

            int firstDataRow = mHeaderRows;
            for (int i = firstDataRow; i < sheet.Rows.Count; i++)
            {
                DataRow row = sheet.Rows[i];
                string ID = row[sheet.Columns[0]].ToString();
                if (ID.Length <= 0)
                    ID = string.Format("row_{0}", i);

                var rowObject = convertRowToDict(sheet, row, lowcase, firstDataRow, excludePrefix, cellJson);
                // 多余的字段
                // rowObject[ID] = ID;
                int idx = i - firstDataRow;
                importData[idx] = rowObject;
            }

            return importData;
        }

        public string GetStringValue(object value)
        {
            if (value == null)
            {
                return string.Empty;
            }
            return value.ToString();
        }

        public static bool GetBoolValue(object value)
        {
            if (value == null)
            {
                return false;
            }
            return System.Convert.ToBoolean(value);
        }

        public long GetLongValue(object value)
        {
            if (value == null)
            {
                return 0;
            }

            if (String.IsNullOrEmpty(value.ToString()))
            {
                return 0;
            }

            return System.Convert.ToInt64(value);
        }

        public int GetIntValue(object value)
        {
            if (value == null)
            {
                return 0;
            }

            if (String.IsNullOrEmpty(value.ToString()))
            {
                return 0;
            }

            return System.Convert.ToInt32(value);
        }

        public Decimal GetFloatValue(object value)
        {
            if (value == null)
            {
                return System.Convert.ToDecimal(0f);
            }
            return System.Convert.ToDecimal(value);
        }

        public static long GetFixedValue(object value)
        {
            if (value == null)
            {
                return 0;
            }
            return (long)(System.Convert.ToDouble(value) * 10000);
        }

        public List<int> GetIntListValue(object o)
        {
            if (o == null)
                return new List<int>();

            string value = o.ToString();
            List<int> intList = new List<int>();
            string[] ints = value.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i < ints.Length; i++)
            {
                try
                {
                    intList.Add(int.Parse(ints[i]));
                }
                catch (Exception e)
                { }
            }
                return intList;
        }

        public List<List<string>> GetListStringListValue(object o)
        {
            if (o == null)
                return new List<List<string>>();

            List<List<string>> listIntList = new List<List<string>>();

            string value = o.ToString();
            string[] intInts = value.TrimStart(new char[] { '[' }).TrimEnd(new char[] { ']' }).Split(new string[] { "][" }, StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i < intInts.Length; i++)
            {
                List<string> intList = new List<string>();
                string[] ints = intInts[i].Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                for (int j = 0; j < ints.Length; j++)
                {
                    intList.Add(ints[j]);
                }
                listIntList.Add(intList);
            }
            return listIntList;
        }

        public List<List<int>> GetListIntListValue(object o)
        {
            if (o == null)
                return new List<List<int>>();

            List<List<int>> listIntList = new List<List<int>>();

            string value = o.ToString();
            string[] intInts = value.TrimStart(new char[] { '[' }).TrimEnd(new char[] { ']' }).Split(new string[] { "][" }, StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i < intInts.Length; i++)
            {
                List<int> intList = new List<int>();
                string[] ints = intInts[i].Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                for (int j = 0; j < ints.Length; j++)
                {
                    intList.Add(int.Parse(ints[j]));
                }
                listIntList.Add(intList);
            }
            return listIntList;
        }

        public List<List<double>> GetListRealFloatListValue(object o)
        {
            if (o == null)
                return new List<List<double>>();

            List<List<double>> listIntList = new List<List<double>>();

            string value = o.ToString();
            string[] intInts = value.TrimStart(new char[] { '[' }).TrimEnd(new char[] { ']' }).Split(new string[] { "][" }, StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i < intInts.Length; i++)
            {
                List<double> intList = new List<double>();
                string[] ints = intInts[i].Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                for (int j = 0; j < ints.Length; j++)
                {
                    intList.Add(System.Convert.ToDouble(ints[j]));
                }
                listIntList.Add(intList);
            }
            return listIntList;
        }

        public List<List<long>> GetListFixedListValue(object o)
        {
            if (o == null)
                return new List<List<long>>();

            List<List<long>> listIntList = new List<List<long>>();

            string value = o.ToString();
            string[] intInts = value.TrimStart(new char[] { '[' }).TrimEnd(new char[] { ']' }).Split(new string[] { "][" }, StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i < intInts.Length; i++)
            {
                List<long> intList = new List<long>();
                string[] ints = intInts[i].Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                for (int j = 0; j < ints.Length; j++)
                {
                    intList.Add((long)(System.Convert.ToDouble(ints[j]) * 10000));
                }
                listIntList.Add(intList);
            }
            return listIntList;
        }

        public List<double> GetFloatListValue(object o)
        {
            if (o == null)
                return new List<double>();

            string value = o.ToString();
            List<double> floatList = new List<double>();
            string[] floats = value.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i < floats.Length; i++)
            {
                floatList.Add(System.Convert.ToDouble(floats[i]));
            }
            return floatList;
        }

        public List<long> GetFixedListValue(object o)
        {
            if (o == null)
                return new List<long>();

            string value = o.ToString();
            List<long> fxiedList = new List<long>();
            string[] fixeds = value.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i < fixeds.Length; i++)
            {
                fxiedList.Add((long)(System.Convert.ToDouble(fixeds[i]) * 10000));
            }
            return fxiedList;
        }

        public List<string> GetStringListValue(object o)
        {
            if (o == null)
                return new List<string>();

            string value = o.ToString();
            List<string> stringList = new List<string>();
            string[] strings = value.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i < strings.Length; i++)
            {
                stringList.Add(strings[i]);
            }
            return stringList;
        }

        public List<string> GetIntPairValue(object value)
        {
            if (value == null)
            {
                return new List<string>();
            }

            string str = value.ToString().Replace("[", "").Replace("]", "");
            List<string> strList = new List<string>(str.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries));
            return strList;
        }


        public object GetObjectValue(object value, object type)
        {
            try
            {
                string fieldType = type.ToString();

                if (fieldType == "string")
                {
                    return GetStringValue(value);
                }
                if (fieldType == "bool")
                {
                    return GetBoolValue(value);
                }
                else if (fieldType == "int" || fieldType == "enum")
                {
                    return GetIntValue(value);
                }
                else if (fieldType == "float")
                {
                    return GetFixedValue(value);
                }
                else if (fieldType == "realFloat")
                {
                    return GetFloatValue(value);
                }
                else if (fieldType == "List<realFloat>")
                {
                    return GetFloatListValue(value);
                }
                else if (fieldType == "List<int>")
                {
                    return GetIntListValue(value);
                }
                else if (fieldType == "List<float>")
                {
                    return GetFixedListValue(value);
                }
                else if (fieldType == "List<string>")
                {
                    return GetStringListValue(value);
                }
                else if (fieldType == "List<List<string>>")
                {
                    return GetListStringListValue(value);
                }
                else if (fieldType == "List<int,int>")
                {
                    return GetIntPairValue(value);
                }
                else if (fieldType == "List<List<int>>")
                {
                    return GetListIntListValue(value);
                }
                else if (fieldType == "List<List<float>>")
                {
                    return GetListFixedListValue(value);
                }
                else if (fieldType == "List<List<realFloat>>")
                {
                    return GetListRealFloatListValue(value);
                }
                else if (fieldType == "long")
                {
                    return GetLongValue(value);
                }

                return null;
            }
            catch (Exception e)
            {
                return null;
            }
        }


        /// <summary>
        /// 把一行数据转换成一个对象，每一列是一个属性
        /// </summary>
        private Dictionary<string, object> convertRowToDict(DataTable sheet, DataRow row, bool lowcase, int firstDataRow, string excludePrefix, bool cellJson)
        {
            var rowData = new Dictionary<string, object>();
            int col = 0;
            for (int i = 0; i < sheet.Columns.Count; i++)
            {
                string columnName = sheet.Rows[1][i].ToString();
                string skip = sheet.Rows[0][i].ToString();
                if (excludePrefix.Length > 0 && columnName.StartsWith(excludePrefix))
                    continue;
                if (skip != "s" && skip != "cs" && skip != "sc")
                    continue;

                //导出类型
                string valueTypestr = sheet.Rows[2][i].ToString();

                object value = row[i];

//                 if (value.GetType() == typeof(System.DBNull))
//                 {
//                     value = getColumnDefault(sheet, i, firstDataRow);
//                 }
//                 else if (value.GetType() == typeof(double))
//                 { // 去掉数值字段的“.0”
//                     double num = (double)value;
//                     if ((int)num == num)
//                         value = (int)num;
//                 }

                // 表头自动转换成小写
                if (lowcase)
                    columnName = columnName.ToLower();

                if (string.IsNullOrEmpty(columnName))
                    columnName = string.Format("col_{0}", col);

                rowData[columnName] = GetObjectValue(value, valueTypestr);
                col++;
            }

            return rowData;
        }

        /// <summary>
        /// 对于表格中的空值，找到一列中的非空值，并构造一个同类型的默认值
        /// </summary>
        private object getColumnDefault(DataTable sheet, int column, int firstDataRow)
        {
            for (int i = firstDataRow; i < sheet.Rows.Count; i++)
            {
                string valueTypest = sheet.Rows[3][column].ToString();
                Type valueType;
                if (valueTypest == "int")
                {
                    return 0;
                }
                else if (valueTypest == "string")
                {
                    return "";
                }
            }
            return "";
        }

        /// <summary>
        /// 将内部数据转换成Json文本，并保存至文件
        /// </summary>
        /// <param name="jsonPath">输出文件路径</param>
        public void SaveToFile(string filePath, Encoding encoding)
        {
            foreach (KeyValuePair<string, object>kvp in mData)
            {
                string strContext = JsonConvert.SerializeObject(kvp.Value, jsonSettings);
                using (FileStream file = new FileStream(filePath+kvp.Key+".json", FileMode.Create, FileAccess.Write))
                {
                    using (TextWriter writer = new StreamWriter(file, encoding))
                        writer.Write(strContext);
                }
            }
        }
    }
}
