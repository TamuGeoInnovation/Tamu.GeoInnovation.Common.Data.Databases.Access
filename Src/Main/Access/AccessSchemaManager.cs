using System;
using System.Data;
using System.Data.Odbc;
using System.Data.OleDb;
using System.Diagnostics;
using System.Runtime.InteropServices;
using ADOX;
using USC.GISResearchLab.Common.Core.Databases;
using USC.GISResearchLab.Common.Databases.QueryManagers;
using USC.GISResearchLab.Common.Databases.SchemaManagers;
using USC.GISResearchLab.Common.Databases.SqlServer;
using USC.GISResearchLab.Common.Databases.TypeConverters;

using USC.GISResearchLab.Common.Utils.Databases.TableDefinitions;

using USC.GISResearchLab.Common.Utils.Files;
using USC.GISResearchLab.Common.Diagnostics.TraceEvents;
using USC.GISResearchLab.Common.Utils.Exceptions;

namespace USC.GISResearchLab.Common.Databases.Access
{
    public class AccessSchemaManager : AbstractSchemaManager
    {

        public AccessSchemaManager(DataProviderType dataProviderType, string connectionString)
        {
            ConnectionString = connectionString;
            this.DatabaseType = DatabaseType.Access;
            this.DataProviderType = dataProviderType;
            QueryManager = new QueryManager(DataProviderType, DatabaseType, ConnectionString);
        }

        public AccessSchemaManager(string pathToDatabaseDlls, DataProviderType dataProviderType, string connectionString)
        {
            ConnectionString = connectionString;
            this.DatabaseType = DatabaseType.Access;
            this.DataProviderType = dataProviderType;
            PathToDatabaseDLLs = pathToDatabaseDlls;
            QueryManager = new QueryManager(pathToDatabaseDlls, DataProviderType, DatabaseType, ConnectionString);
        }

        public override void CreateDatabase()
        {
            try
            {
                string databasePath = "";
                switch (DataProviderType)
                {
                    case DataProviderType.Odbc:
                        databasePath = ((OdbcConnection)Connection).DataSource;
                        break;
                    case DataProviderType.OleDb:
                        databasePath = ((OleDbConnection)Connection).DataSource;
                        break;
                    default:
                        throw new Exception("Unexpected DataProviderType: " + DataProviderType);
                }

                Catalog cat = new Catalog();
                var o = cat.Create("Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + databasePath + ";" + "Jet OLEDB:Engine Type=5");

                Marshal.ReleaseComObject(cat);
                cat = null;
                GC.Collect();
            }
            catch (Exception ex)
            {
                string msg = "Error creating database: " + ex.Message;
                throw new Exception(msg, ex);
            }
        }

        public override TableDefinition GetTableDefinition(string table)
        {
            return null;
        }

        public override TableColumn[] GetColumns(string table)
        {
            return null;
        }

        public override TableColumn[] GetColumns(string table, bool shouldOpenAndClose)
        {
            return null;
        }

        public override string BuildCreateTableStatement(TableDefinition tableDefinition)
        {
            string ret = "";
            if (tableDefinition != null)
            {
                string columnsSql = "";
                string keysSql = "";

                if (tableDefinition.TableColumns != null)
                    for (int i = 0; i < tableDefinition.TableColumns.Length; i++)
                    {

                        TableColumn column = tableDefinition.TableColumns[i];
                        if (i != 0)
                        {
                            columnsSql += ", ";
                        }

                        columnsSql += " [" + column.Name + "]";
                        columnsSql += " " + DatabaseTypeConverter.GetTypeAsString(column.DatabaseSuperDataType);
                        if (column.Length > 0)
                        {
                            columnsSql += "(" + column.Length + ")";
                        }
                        else
                        {
                            int defaultLength = DatabaseTypeConverter.GetDefaultLength(column.DatabaseSuperDataType);
                            int defaultPrecision = DatabaseTypeConverter.GetDefaultPrecision(column.DatabaseSuperDataType);

                            if (defaultLength > 0)
                            {
                                columnsSql += "(";
                                columnsSql += defaultLength;

                                if (defaultPrecision > 0)
                                {
                                    columnsSql += ",";
                                    columnsSql += defaultPrecision;
                                }

                                columnsSql += ")";
                            }
                        }



                        if (!column.IsNullable)
                        {
                            columnsSql += " NOT NULL ";
                        }

                        if (column.IsAutoIncrement)
                        {
                            columnsSql += " COUNTER ";
                        }

                        if (column.DefaultValue != null)
                        {
                            columnsSql += " " + column.DefaultValue.ToString();
                        }

                        if (column.IsPrimaryKey)
                        {
                            if (keysSql == String.Empty)
                            {
                                keysSql += " PRIMARY KEY (";
                            }
                            else
                            {
                                keysSql += ", ";
                            }

                            keysSql += " [" + column.Name + "]";
                        }
                    }

                ret = " CREATE TABLE [" + tableDefinition.Name + "]";

                if (columnsSql != String.Empty)
                {
                    if (keysSql != String.Empty)
                    {
                        ret += " ( " + columnsSql + ", " + keysSql + ") )";
                    }
                    else
                    {
                        ret += " ( " + columnsSql + "  )";
                    }
                }
            }
            return ret;
        }

        public override void RemoveTableFromDatabase(string tableName)
        {
            try
            {
                string sql = "DELETE TABLE [" + tableName + "]";
                QueryManager.ExecuteNonQuery(CommandType.Text, sql, true);
            }
            catch (Exception ex)
            {
                string msg = "Error removing table: " + ex.Message;
                throw new Exception(msg, ex);
            }
        }

        public override void RemoveIndexFromTable(string tableName, string indexName)
        {
            throw new NotImplementedException();
        }

        public override void RemoveSpatialIndexFromTable(string tableName, string indexName)
        {
            throw new NotImplementedException();
        }

        public override void AddColumnsToTable(string tableName, string[] columnNames, DatabaseSuperDataType[] dataTypes)
        {
            for (int i = 0; i < columnNames.Length; i++)
            {
                AddColumnToTable(tableName, columnNames[i], dataTypes[i]);
            }
        }

        public override void AddColumnToTable(string tableName, string columnName, DatabaseSuperDataType dataType)
        {
            AccessTypeConverter typeConverter = new AccessTypeConverter();
            int defaultLength = typeConverter.GetDefaultLength(dataType);
            int defaultPrecision = typeConverter.GetDefaultPrecision(dataType);
            AddColumnToTable(tableName, columnName, dataType, false, defaultLength, defaultPrecision);
        }

        public override void AddColumnToTable(string tableName, string columnName, DatabaseSuperDataType dataType, bool nullable, int maxLength, int precision)
        {
            try
            {
                string dbTypeName = DatabaseTypeConverter.GetTypeAsString(dataType);

                string sql = "";
                sql += " ALTER TABLE ";

                if (!tableName.Trim().StartsWith("["))
                {
                    sql += " [";
                }

                sql += tableName;

                if (!tableName.Trim().EndsWith("]"))
                {
                    sql += "]";
                }

                sql += " ADD ";

                if (!columnName.Trim().StartsWith("["))
                {
                    sql += " [";
                }

                sql += columnName;

                if (!columnName.Trim().EndsWith("]"))
                {
                    sql += "]";
                }

                sql += " " + dbTypeName;

                if (maxLength > 0 || precision > 0)
                {
                    sql += " (";
                    if (maxLength > 0)
                    {
                        sql += maxLength;
                    }

                    if (precision > 0)
                    {
                        sql += " , ";
                        sql += precision;
                    }

                    sql += " )";
                }

                if (!nullable)
                {
                    sql += " NOT NULL ";
                }


                QueryManager.ExecuteNonQuery(CommandType.Text, sql, true);
            }
            catch (Exception ex)
            {
                string msg = "Error adding column to table: " + ex.Message;
                throw new Exception(msg, ex);
            }
        }

        public override string[] GetDatabases()
        {
            return null;
        }

        public override string[] GetDatabases(bool shouldOpenAndClose)
        {
            return null;
        }

        public override string[] GetDatabases(DatabaseNameListingOptions opt)
        {
            throw new NotImplementedException();
        }

        public override string[] GetDatabases(bool shouldOpenAndClose, DatabaseNameListingOptions opt)
        {
            throw new NotImplementedException();
        }

        public override DataTable GetDatabasesAsDataTable()
        {
            throw new NotImplementedException();
        }

        public override DataTable GetDatabasesAsDataTable(bool shouldOpenAndClose)
        {
            throw new NotImplementedException();
        }

        public override DataTable GetDatabasesAsDataTable(bool shouldOpenAndClose, DatabaseNameListingOptions opt)
        {
            throw new NotImplementedException();
        }

        public override string[] GetTables()
        {
            return null;
        }

        public override string[] GetTables(bool shouldOpenAndClose)
        {
            return null;
        }

        public override DataTable GetTablesAsDataTable()
        {
            return null;
        }

        public override DataTable GetTablesAsDataTable(bool shouldOpenAndClose)
        {
            return null;
        }

        public override string[] GetTableIndexes(string tableName)
        {
            throw new NotImplementedException();
        }

        public override string[] GetTableIndexes(string tableName, bool shouldOpenAndClose)
        {
            throw new NotImplementedException();
        }

        public override string[] GetTableSpatialIndexes(string tableName)
        {
            throw new NotImplementedException();
        }

        public override string[] GetTableSpatialIndexes(string tableName, bool shouldOpenAndClose)
        {
            throw new NotImplementedException();
        }

        public override string GetTableClusteredIndex(string tableName)
        {
            throw new NotImplementedException();
        }

        public override string GetTableClusteredIndex(string tableName, bool shouldOpenAndClose)
        {
            throw new NotImplementedException();
        }

        public override void RemoveConstraintFromTable(string tableName, string constraintName)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// MBD compact method (c) 2004 Alexander Youmashev
        /// !!IMPORTANT!!
        /// !make sure there's no open connections to your db before calling this method!
        /// !!IMPORTANT!!
        /// </summary>
        /// <param name="connectionString">connection string to your db</param>
        /// <param name="mdwfilename">FULL name of an MDB file you want to compress.</param>
        ///     <see cref="http://www.codeproject.com/Articles/7775/Compact-and-Repair-Access-Database-using-C-and-lat"/>
        public static bool CompactAccessDB(string mdbFileName, TraceSource trace)
        {
            object[] oParams;
            bool ret = false;
            object objJRO = null;
            string tempFilepath = string.Empty;
            try
            {
                do
                {
                    tempFilepath = System.IO.Path.GetTempPath() + System.IO.Path.GetRandomFileName() + FileUtils.GetExtension(mdbFileName, true);
                } while (System.IO.File.Exists(tempFilepath));

                //create an inctance of a Jet Replication Object
                objJRO = Activator.CreateInstance(Type.GetTypeFromProgID("JRO.JetEngine"));

                //filling Parameters array
                //cnahge "Jet OLEDB:Engine Type=5" to an appropriate value
                // or leave it as is if you db is JET4X format (access 2000,2002)
                //(yes, jetengine5 is for JET4X, no misprint here)
                oParams = new object[] { 
                    "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + mdbFileName + ";Jet OLEDB:Engine Type=5",
                    "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + tempFilepath + ";Jet OLEDB:Engine Type=5" };

                //invoke a CompactDatabase method of a JRO object
                //pass Parameters array
                objJRO.GetType().InvokeMember("CompactDatabase", System.Reflection.BindingFlags.InvokeMethod, null, objJRO, oParams);

                //database is compacted now
                //to a new file C:\tempdb.mdw
                //let's copy it over an old one and delete it

                System.IO.File.Delete(mdbFileName);
                System.IO.File.Move(tempFilepath, mdbFileName);
                ret = true;
            }
            catch (Exception ex)
            {
                if (trace != null)
                    trace.TraceEvent(TraceEventType.Error, (int)ExceptionEvents.ExceptionOccurred, "CompactAccessDB > " + ExceptionUtils.PrepareErrorMessage2(ex));
            }
            finally
            {
                //clean up (just in case)
                FileUtils.DeleteFile(tempFilepath);
                if (objJRO != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objJRO);
                    objJRO = null;
                }
            }
            return ret;
        }

        public override void AddGeogIndexToDatabase(string tableName)
        {
            throw new NotImplementedException();
        }

        public override void AddGeogIndexToDatabase(string tableName, bool shouldOpenCloseConnection)
        {
            throw new NotImplementedException();
        }
    }
}
