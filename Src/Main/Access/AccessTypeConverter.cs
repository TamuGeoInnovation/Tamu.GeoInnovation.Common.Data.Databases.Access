using System;
using System.Data.Odbc;
using System.Data.OleDb;
using USC.GISResearchLab.Common.Core.Databases;
using USC.GISResearchLab.Common.Databases.TypeConverters;
using USC.GISResearchLab.Common.Databases.TypeConverters.DatabaseTypeConverters;

namespace USC.GISResearchLab.Common.Databases.Access
{
    public class AccessTypeConverter : AbstractDatabaseTypeConverterManager
    {

        #region TypeNames

        public static string TYPENAME_Binary = "Binary";
        public static string TYPENAME_Bit = "Bit";
        public static string TYPENAME_Counter = "Counter";
        public static string TYPENAME_Currency = "Currency";
        public static string TYPENAME_DateTime = "DateTime";
        public static string TYPENAME_Decimal = "Decimal";
        public static string TYPENAME_Double = "Double";
        public static string TYPENAME_Guid = "Guid";
        public static string TYPENAME_Long = "Long";
        public static string TYPENAME_LongBinary = "LongBinary";
        public static string TYPENAME_LongText = "LongText";
        public static string TYPENAME_Numeric = "Numeric";
        public static string TYPENAME_Single = "Single";
        public static string TYPENAME_Short = "Short";
        public static string TYPENAME_UByte = "Unsigned Byte";
        public static string TYPENAME_VarChar = "VarChar";
        public static string TYPENAME_VarBinary = "VarBinary";
        #endregion

        public AccessTypeConverter()
        {
            DatabaseType = DatabaseType.Access;
        }


        public override int GetDefaultLength(DatabaseSuperDataType type)
        {
            int ret = -1;

            switch (type)
            {

                case DatabaseSuperDataType.Decimal:
                case DatabaseSuperDataType.Double:
                case DatabaseSuperDataType.Float:
                    ret = 28;
                    break;
                case DatabaseSuperDataType.VarChar:
                    ret = 255;
                    break;
                default:
                    break;
            }
            return ret;
        }


        public override int GetDefaultPrecision(DatabaseSuperDataType type)
        {
            int ret = -1;

            switch (type)
            {
                case DatabaseSuperDataType.Decimal:
                case DatabaseSuperDataType.Double:
                case DatabaseSuperDataType.Float:
                    ret = 18;
                    break;
                default:
                    break;
            }
            return ret;
        }

        public override object ConvertType(object dbType, DatabaseType databaseType)
        {
            throw new NotImplementedException();
        }

        public DatabaseSuperDataType ToSuperType(object dbType)
        {
            DatabaseSuperDataType ret = DatabaseSuperDataType.VarChar;
            if (dbType.GetType() == typeof(OleDbType))
            {
                ret = ToSuperType((OleDbType)dbType);
            }
            else if (dbType.GetType() == typeof(OdbcType))
            {
                ret = ToSuperType((OdbcType)dbType);
            }
            return ret;
        }

        public DatabaseSuperDataType ToSuperType(OleDbType type)
        {
            DatabaseSuperDataType ret;

            switch (type)
            {
                case OleDbType.BigInt:
                    ret = DatabaseSuperDataType.Long;
                    break;
                case OleDbType.Binary:
                    ret = DatabaseSuperDataType.Binary;
                    break;
                case OleDbType.Boolean:
                    ret = DatabaseSuperDataType.Bit;
                    break;
                case OleDbType.BSTR:
                    ret = DatabaseSuperDataType.VarChar;
                    break;
                case OleDbType.Char:
                    ret = DatabaseSuperDataType.VarChar;
                    break;
                case OleDbType.Currency:
                    ret = DatabaseSuperDataType.Currency;
                    break;
                case OleDbType.Date:
                    ret = DatabaseSuperDataType.DateTime;
                    break;
                case OleDbType.DBDate:
                    ret = DatabaseSuperDataType.DateTime;
                    break;
                case OleDbType.DBTime:
                    ret = DatabaseSuperDataType.DateTime;
                    break;
                case OleDbType.DBTimeStamp:
                    ret = DatabaseSuperDataType.DateTime;
                    break;
                case OleDbType.Decimal:
                    ret = DatabaseSuperDataType.Numeric;
                    break;
                case OleDbType.Double:
                    ret = DatabaseSuperDataType.Double;
                    break;
                case OleDbType.Empty:
                    ret = DatabaseSuperDataType.VarChar;
                    break;
                case OleDbType.Error:
                    ret = DatabaseSuperDataType.VarChar;
                    break;
                case OleDbType.Filetime:
                    ret = DatabaseSuperDataType.DateTime;
                    break;
                case OleDbType.Guid:
                    ret = DatabaseSuperDataType.Guid;
                    break;
                case OleDbType.IDispatch:
                    ret = DatabaseSuperDataType.VarChar;
                    break;
                case OleDbType.Integer:
                    ret = DatabaseSuperDataType.Numeric;
                    break;
                case OleDbType.IUnknown:
                    ret = DatabaseSuperDataType.VarChar;
                    break;
                case OleDbType.LongVarBinary:
                    ret = DatabaseSuperDataType.VarBinary;
                    break;
                case OleDbType.LongVarChar:
                    ret = DatabaseSuperDataType.VarChar;
                    break;
                case OleDbType.LongVarWChar:
                    ret = DatabaseSuperDataType.VarChar;
                    break;
                case OleDbType.Numeric:
                    ret = DatabaseSuperDataType.Numeric;
                    break;
                case OleDbType.PropVariant:
                    ret = DatabaseSuperDataType.VarChar;
                    break;
                case OleDbType.Single:
                    ret = DatabaseSuperDataType.Single;
                    break;
                case OleDbType.SmallInt:
                    ret = DatabaseSuperDataType.Short;
                    break;
                case OleDbType.TinyInt:
                    ret = DatabaseSuperDataType.Short;
                    break;
                case OleDbType.UnsignedBigInt:
                    ret = DatabaseSuperDataType.UByte;
                    break;
                case OleDbType.UnsignedInt:
                    ret = DatabaseSuperDataType.UByte;
                    break;
                case OleDbType.UnsignedSmallInt:
                    ret = DatabaseSuperDataType.UByte;
                    break;
                case OleDbType.UnsignedTinyInt:
                    ret = DatabaseSuperDataType.UByte;
                    break;
                case OleDbType.VarBinary:
                    ret = DatabaseSuperDataType.VarBinary;
                    break;
                case OleDbType.VarChar:
                    ret = DatabaseSuperDataType.VarChar;
                    break;
                case OleDbType.Variant:
                    ret = DatabaseSuperDataType.VarChar;
                    break;
                case OleDbType.VarNumeric:
                    ret = DatabaseSuperDataType.Numeric;
                    break;
                case OleDbType.VarWChar:
                    ret = DatabaseSuperDataType.UByte;
                    break;
                case OleDbType.WChar:
                    ret = DatabaseSuperDataType.UByte;
                    break;
                default:
                    throw new Exception("Unexpected type: " + type);
            }
            return ret;
        }

        public DatabaseSuperDataType ToSuperType(OdbcType type)
        {
            DatabaseSuperDataType ret;
            switch (type)
            {
                case OdbcType.BigInt:
                    ret = DatabaseSuperDataType.Long;
                    break;
                case OdbcType.Binary:
                    ret = DatabaseSuperDataType.Binary;
                    break;
                case OdbcType.Bit:
                    ret = DatabaseSuperDataType.Bit;
                    break;
                case OdbcType.Char:
                    ret = DatabaseSuperDataType.VarChar;
                    break;
                case OdbcType.Date:
                    ret = DatabaseSuperDataType.DateTime;
                    break;
                case OdbcType.DateTime:
                    ret = DatabaseSuperDataType.DateTime;
                    break;
                case OdbcType.Decimal:
                    ret = DatabaseSuperDataType.Numeric;
                    break;
                case OdbcType.Double:
                    ret = DatabaseSuperDataType.Double;
                    break;
                case OdbcType.Image:
                    ret = DatabaseSuperDataType.Binary;
                    break;
                case OdbcType.Int:
                    ret = DatabaseSuperDataType.Numeric;
                    break;
                case OdbcType.NChar:
                    ret = DatabaseSuperDataType.VarChar;
                    break;
                case OdbcType.NText:
                    ret = DatabaseSuperDataType.VarChar;
                    break;
                case OdbcType.Numeric:
                    ret = DatabaseSuperDataType.Numeric;
                    break;
                case OdbcType.NVarChar:
                    ret = DatabaseSuperDataType.NVarChar;
                    break;
                case OdbcType.Real:
                    ret = DatabaseSuperDataType.Numeric;
                    break;
                case OdbcType.SmallDateTime:
                    ret = DatabaseSuperDataType.DateTime;
                    break;
                case OdbcType.SmallInt:
                    ret = DatabaseSuperDataType.Short;
                    break;
                case OdbcType.Text:
                    ret = DatabaseSuperDataType.VarChar;
                    break;
                case OdbcType.Time:
                    ret = DatabaseSuperDataType.DateTime;
                    break;
                case OdbcType.Timestamp:
                    ret = DatabaseSuperDataType.DateTime;
                    break;
                case OdbcType.TinyInt:
                    ret = DatabaseSuperDataType.Short;
                    break;
                case OdbcType.UniqueIdentifier:
                    ret = DatabaseSuperDataType.Guid;
                    break;
                case OdbcType.VarBinary:
                    ret = DatabaseSuperDataType.VarBinary;
                    break;
                case OdbcType.VarChar:
                    ret = DatabaseSuperDataType.VarChar;
                    break;
                default:
                    throw new Exception("Unexpected type: " + type);
            }
            return ret;
        }

        public override string GetTypeAsString(DatabaseSuperDataType type)
        {

            string ret = null;

            switch (type)
            {
                case DatabaseSuperDataType.BigInt:
                    ret = TYPENAME_Numeric;
                    break;
                case DatabaseSuperDataType.Binary:
                    ret = TYPENAME_Binary;
                    break;
                case DatabaseSuperDataType.Bit:
                    ret = TYPENAME_Bit;
                    break;
                case DatabaseSuperDataType.Blob:
                    ret = TYPENAME_Binary;
                    break;
                case DatabaseSuperDataType.Boolean:
                    ret = TYPENAME_Bit;
                    break;
                case DatabaseSuperDataType.BSTR:
                    ret = TYPENAME_VarChar;
                    break;
                case DatabaseSuperDataType.Char:
                    ret = TYPENAME_VarChar;
                    break;
                case DatabaseSuperDataType.Counter:
                    ret = TYPENAME_Counter;
                    break;
                case DatabaseSuperDataType.Currency:
                    ret = TYPENAME_Currency;
                    break;
                case DatabaseSuperDataType.Date:
                    ret = TYPENAME_DateTime;
                    break;
                case DatabaseSuperDataType.DateTime:
                    ret = TYPENAME_DateTime;
                    break;
                case DatabaseSuperDataType.DateTime2:
                    ret = TYPENAME_DateTime;
                    break;
                case DatabaseSuperDataType.DateTimeOffset:
                    ret = TYPENAME_DateTime;
                    break;
                case DatabaseSuperDataType.DBDate:
                    ret = TYPENAME_DateTime;
                    break;
                case DatabaseSuperDataType.DBTime:
                    ret = TYPENAME_DateTime;
                    break;
                case DatabaseSuperDataType.DBTimeStamp:
                    ret = TYPENAME_DateTime;
                    break;
                case DatabaseSuperDataType.Decimal:
                    ret = TYPENAME_Decimal;
                    break;
                case DatabaseSuperDataType.Double:
                    ret = TYPENAME_Decimal;
                    break;
                case DatabaseSuperDataType.Empty:
                    ret = TYPENAME_VarChar;
                    break;
                case DatabaseSuperDataType.Error:
                    ret = TYPENAME_VarChar;
                    break;
                case DatabaseSuperDataType.Filetime:
                    ret = TYPENAME_DateTime;
                    break;
                case DatabaseSuperDataType.Float:
                    ret = TYPENAME_Decimal;
                    break;
                case DatabaseSuperDataType.Geometry:
                    ret = TYPENAME_VarChar;
                    break;
                case DatabaseSuperDataType.Guid:
                    ret = TYPENAME_Guid;
                    break;
                case DatabaseSuperDataType.IDispatch:
                    ret = TYPENAME_VarChar;
                    break;
                case DatabaseSuperDataType.Image:
                    ret = TYPENAME_VarBinary;
                    break;
                case DatabaseSuperDataType.Int16:
                    ret = TYPENAME_Numeric;
                    break;
                case DatabaseSuperDataType.Int24:
                    ret = TYPENAME_Numeric;
                    break;
                case DatabaseSuperDataType.Int32:
                    ret = TYPENAME_Long;
                    break;
                case DatabaseSuperDataType.Int64:
                    ret = TYPENAME_Numeric;
                    break;
                case DatabaseSuperDataType.IUnknown:
                    ret = TYPENAME_VarChar;
                    break;
                case DatabaseSuperDataType.Long:
                    ret = TYPENAME_Long;
                    break;
                case DatabaseSuperDataType.LongBinary:
                    ret = TYPENAME_LongBinary;
                    break;
                case DatabaseSuperDataType.LongText:
                    ret = TYPENAME_LongText;
                    break;
                case DatabaseSuperDataType.LongVarBinary:
                    ret = TYPENAME_LongBinary;
                    break;
                case DatabaseSuperDataType.LongVarChar:
                    ret = TYPENAME_LongText;
                    break;
                case DatabaseSuperDataType.LongVarWChar:
                    ret = TYPENAME_LongText;
                    break;
                case DatabaseSuperDataType.MediumBlob:
                    ret = TYPENAME_VarBinary;
                    break;
                case DatabaseSuperDataType.Money:
                    ret = TYPENAME_Currency;
                    break;
                case DatabaseSuperDataType.NChar:
                    ret = TYPENAME_VarChar;
                    break;
                case DatabaseSuperDataType.Newdate:
                    ret = TYPENAME_DateTime;
                    break;
                case DatabaseSuperDataType.NewDecimal:
                    ret = TYPENAME_Decimal;
                    break;
                case DatabaseSuperDataType.NText:
                    ret = TYPENAME_VarChar;
                    break;
                case DatabaseSuperDataType.Numeric:
                    ret = TYPENAME_Decimal;
                    break;
                case DatabaseSuperDataType.NVarChar:
                    ret = TYPENAME_VarChar;
                    break;
                case DatabaseSuperDataType.PropVariant:
                    ret = TYPENAME_VarChar;
                    break;
                case DatabaseSuperDataType.Real:
                    ret = TYPENAME_Decimal;
                    break;
                case DatabaseSuperDataType.Set:
                    ret = TYPENAME_Binary;
                    break;
                case DatabaseSuperDataType.Short:
                    ret = TYPENAME_Short;
                    break;
                case DatabaseSuperDataType.Single:
                    ret = TYPENAME_Single;
                    break;
                case DatabaseSuperDataType.SmallDateTime:
                    ret = TYPENAME_DateTime;
                    break;
                case DatabaseSuperDataType.SmallInt:
                    ret = TYPENAME_Short;
                    break;
                case DatabaseSuperDataType.SmallMoney:
                    ret = TYPENAME_Currency;
                    break;
                case DatabaseSuperDataType.String:
                    ret = TYPENAME_VarChar;
                    break;
                case DatabaseSuperDataType.Structured:
                    ret = TYPENAME_VarChar;
                    break;
                case DatabaseSuperDataType.Text:
                    ret = TYPENAME_VarChar;
                    break;
                case DatabaseSuperDataType.Time:
                    ret = TYPENAME_DateTime;
                    break;
                case DatabaseSuperDataType.Timestamp:
                    ret = TYPENAME_VarChar;
                    break;
                case DatabaseSuperDataType.TinyBlob:
                    ret = TYPENAME_VarBinary;
                    break;
                case DatabaseSuperDataType.TinyInt:
                    ret = TYPENAME_Short;
                    break;
                case DatabaseSuperDataType.TinyText:
                    ret = TYPENAME_VarChar;
                    break;
                case DatabaseSuperDataType.UByte:
                    ret = TYPENAME_UByte;
                    break;
                case DatabaseSuperDataType.Udt:
                    ret = TYPENAME_VarChar;
                    break;
                case DatabaseSuperDataType.UInt16:
                    ret = TYPENAME_UByte;
                    break;
                case DatabaseSuperDataType.UInt24:
                    ret = TYPENAME_UByte;
                    break;
                case DatabaseSuperDataType.UInt32:
                    ret = TYPENAME_UByte;
                    break;
                case DatabaseSuperDataType.UInt64:
                    ret = TYPENAME_UByte;
                    break;
                case DatabaseSuperDataType.UniqueIdentifier:
                    ret = TYPENAME_Guid;
                    break;
                case DatabaseSuperDataType.UnsignedBigInt:
                    ret = TYPENAME_UByte;
                    break;
                case DatabaseSuperDataType.UnsignedInt:
                    ret = TYPENAME_UByte;
                    break;
                case DatabaseSuperDataType.UnsignedSmallInt:
                    ret = TYPENAME_UByte;
                    break;
                case DatabaseSuperDataType.UnsignedTinyInt:
                    ret = TYPENAME_UByte;
                    break;
                case DatabaseSuperDataType.VarBinary:
                    ret = TYPENAME_VarBinary;
                    break;
                case DatabaseSuperDataType.VarChar:
                    ret = TYPENAME_VarChar;
                    break;
                case DatabaseSuperDataType.Variant:
                    ret = TYPENAME_Binary;
                    break;
                case DatabaseSuperDataType.VarNumeric:
                    ret = TYPENAME_Decimal;
                    break;
                case DatabaseSuperDataType.VarString:
                    ret = TYPENAME_VarChar;
                    break;
                case DatabaseSuperDataType.VarWChar:
                    ret = TYPENAME_VarChar;
                    break;
                case DatabaseSuperDataType.WChar:
                    ret = TYPENAME_VarChar;
                    break;
                case DatabaseSuperDataType.Xml:
                    ret = TYPENAME_VarChar;
                    break;
                case DatabaseSuperDataType.Year:
                    ret = TYPENAME_Decimal;
                    break;
                default:
                    throw new Exception("Unexpected or unimplemented DatabaseSuperDataType: " + type);
            }
            return ret;
        }

    }
}
