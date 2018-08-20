using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Reflection.Emit;
using System.Text;
using System.Threading.Tasks;

namespace SimpleOrm
{
    public static class SimpleOrm
    {
        static SimpleOrm()
        {
            var typeName = "DynamicType";
            var an = new AssemblyName(typeName);
            AssemblyBuilder ab = AppDomain.CurrentDomain.DefineDynamicAssembly(an, AssemblyBuilderAccess.Run);
            _mb = ab.DefineDynamicModule("Dynamic Module");
        }

        /// <summary>
        /// 从data中读取数据，转化为动态类型的数据
        /// 异常：
        ///     ArgumentNullException :data为null
        ///     InvalidOperationException:can't find table from data
        ///     
        /// </summary>
        /// <param name="data"></param>
        /// <param name="tableName"></param>
        /// <returns></returns>
        public static IEnumerable<dynamic> Read(this DataSet data, string tableName)
        {
            if (data == null)
            {
                throw new ArgumentNullException("Can't read, param 'data' is null");
            }
            DataTable table;
            if (!TryGetTargetTable(data, tableName, out table))
            {
                throw new InvalidOperationException($"can't find table {tableName} from given Datatable 'data'");
            }
            return Read(table);
        }

        /// <summary>
        /// 从data中读取数据，转化为动态类型的数据
        /// 异常：
        ///     ArgumentNullException :table为null
        /// </summary>
        /// <param name="table"></param>
        /// <returns></returns>
        public static IEnumerable<dynamic> Read(this DataTable table)
        {
            if (table == null || table.Columns.Count == 0 || table.Rows.Count == 0)
            {
                throw new ArgumentNullException("Can't read, param 'table' is null or empty");
            }
            var deserializer = BuildDeserializer(GetNameColumnPair(table));
            foreach (DataRow row in table.Rows)
            {
                yield return deserializer(row);
            }
        }

        /// <summary>
        /// 将 list 转换为 DataTable
        /// </summary>
        /// <param name="list"></param>
        /// <returns></returns>
        public static DataTable BuildDataTable<T>(List<T> list)
        {
            if (list == null || !list.Any())
            {
                throw new ArgumentNullException("argument list is null or empty");
            }
            var properties = list[0].GetType().GetProperties().OrderBy(p => p.Name);
            var columns = properties.Select(p => new DataColumn(p.Name, p.PropertyType)).OrderBy(p => p.ColumnName);
            var table = new DataTable();
            table.TableName = typeof(T).Name;
            foreach (var column in columns)
            {
                table.Columns.Add(column);
            }
            foreach (var item in list)
            {
                DataRow row = table.NewRow();
                var values = GetValues(item, properties);
                for (int i = 0; i < values.Length; i++)
                {
                    row[i] = values[i];
                }
                table.Rows.Add(row);
            }
            return table;
        }

        private static object[] GetArray<T>(IList<T> list) {
            var array = new object[list.Count()];
            for (int i = 0; i < list.Count(); i++)
            {
                array[i] = list[i];
            }
            return array;
        }

        private static bool TryGetTargetTable(DataSet data, string tableName, out DataTable table)
        {
            foreach (DataTable item in data.Tables)
            {
                if (item.TableName == tableName)
                {
                    table = item;
                    return true;
                }
            }
            table = null;
            return false;
        }

        private static IEnumerable<string> GetTableColumnsName(DataTable table)
        {
            foreach (DataColumn column in table.Columns)
            {
                yield return column.ColumnName;
            }
        }

        private static IEnumerable<KeyValuePair<string, DataColumn>> GetNameColumnPair(DataTable table)
        {
            foreach (DataColumn column in table.Columns)
            {
                yield return new KeyValuePair<string, DataColumn>(column.ColumnName, column);
            }
        }

        private static Func<DataRow, dynamic> BuildDeserializer(IEnumerable<KeyValuePair<string, DataColumn>> pairs)
        {
            var type = BuildType(pairs);
            return new Func<DataRow, dynamic>(row =>
            {
                object o = Activator.CreateInstance(type);
                foreach (var p in pairs)
                {
                    var property = type.GetField(p.Key);
                    property.SetValue(o, row[p.Value]);
                }
                return o;
            });
        }

        static ModuleBuilder _mb;
        static int _typeVersion = 0;
        private static Type BuildType(IEnumerable<KeyValuePair<string, DataColumn>> pairs)
        {
            var name = "type" + _typeVersion;
            _typeVersion += 1;
            var tb = _mb.DefineType(name, TypeAttributes.Public);
            foreach (var pair in pairs)
            {
                tb.DefineField(pair.Key, pair.Value.DataType, FieldAttributes.Public);
            }
            return tb.CreateType();
        }

        private static object[] GetValues(object o, IEnumerable<PropertyInfo> properties)
        {
            var values = new List<object>();
            foreach (var p in properties)
            {
                values.Add(p.GetValue(o).ToString());
            }
            return values.ToArray();
        }
    }
}
