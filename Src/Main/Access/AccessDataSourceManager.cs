using ADOX;
using System;
using System.IO;
using System.Runtime.InteropServices;
using USC.GISResearchLab.Common.Core.Databases;
using USC.GISResearchLab.Common.Databases.DataSources;
using USC.GISResearchLab.Common.Utils.Files;

namespace USC.GISResearchLab.Common.Databases.Access
{
    public class AccessDataSourceManager : AbstractDataSourceManager
    {

        public AccessDataSourceManager()
        {

        }

        public AccessDataSourceManager(string location, string defualtDatabase, string userName, string password, string[] parameters)
        {
            Location = location;
            DefaultDatabase = defualtDatabase;
            UserName = userName;
            Password = password;
            Parameters = parameters;
        }

        public override void CreateDatabase(DatabaseType databaseType, string databaseName)
        {
            try
            {
                Catalog cat = new Catalog();
                cat.Create("Provider=Microsoft.Jet.OLEDB.4.0;" +
                    "Data Source=" + Path.Combine(Location, databaseName) + ";" +
                    "Jet OLEDB:Engine Type=5");

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

        public override bool Validate(DatabaseType databaseType, string databaseName)
        {
            return FileUtils.FileExists(Path.Combine(Location, databaseName));
        }
    }
}
