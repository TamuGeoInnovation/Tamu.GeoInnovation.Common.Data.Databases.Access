using System;
using System.IO;
using USC.GISResearchLab.Common.Core.Databases;
using USC.GISResearchLab.Common.Databases.ConnectionStringManagers;
using USC.GISResearchLab.Common.Databases.Odbc;
using USC.GISResearchLab.Common.Databases.OleDb;

namespace USC.GISResearchLab.Common.Databases.Access
{
    public class AccessConnectionStringManager : AbstractConnectionStringManager, IConnectionStringManager
    {
        public AccessConnectionStringManager()
        {
            DatabaseType = DatabaseType.Access;
        }

        public AccessConnectionStringManager(string location, string defualtDatabase, string userName, string password, string[] parameters)
        {
            Location = location;
            DefaultDatabase = defualtDatabase;
            UserName = userName;
            Password = password;
            Parameters = parameters;
        }

        public AccessConnectionStringManager(string pathToDatabaseDlls, string location, string defualtDatabase, string userName, string password, string[] parameters)
        {
            PathToDatabaseDLLs = pathToDatabaseDlls;
            Location = location;
            DefaultDatabase = defualtDatabase;
            UserName = userName;
            Password = password;
            Parameters = parameters;
        }

        public override string GetConnectionString(DataProviderType dataProviderType)
        {
            string ret = null;
            switch (dataProviderType)
            {
                case DataProviderType.Odbc:
                    ret = "Driver={" + Drivers.Access + "};DBQ=" + Path.Combine(Location, DefaultDatabase) + ";UID=" + UserName + ";PWD=" + Password + ";";
                    break;
                case DataProviderType.OleDb:
                    ret = "Provider=" + Providers.Access2007 + ";Data Source=" + Path.Combine(Location, DefaultDatabase);
                    break;
                default:
                    throw new Exception("Unexpected dataProviderType: " + dataProviderType);
            }
            return ret;
        }
    }
}
