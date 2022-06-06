using Microsoft.Extensions.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;


// get data from config

var config = new ConfigurationBuilder()
    .AddJsonFile("appsettings.json", optional: false)
    .Build();

string migrateUsersFileName = config.GetSection("MigrateUsersFile")["FileName"];
string migrateUsersSheetName = config.GetSection("MigrateUsersFile")["SheetName"];
string migrateUsersOldLoginColumn = config.GetSection("MigrateUsersFile")["OldLoginColumn"];
string migrateUsersNewLoginColumn = config.GetSection("MigrateUsersFile")["NewLoginColumn"];

string databaseUsersTable = config.GetSection("Database")["UsersTable"];
string databaseUsersLoginColumn = config.GetSection("Database")["UsersLoginColumn"];

string databaseConnectionString = config.GetSection("ConnectionStrings")["DatabaseConnectionString"];


// get login mapping from excel

var fileName = $"{ Directory.GetCurrentDirectory() }\\{ migrateUsersFileName }";
var excelConnectionString = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=Excel 12.0;", fileName);

var adapter = new OleDbDataAdapter($"SELECT * FROM [{ migrateUsersSheetName }$]", excelConnectionString);
var ds = new DataSet();

adapter.Fill(ds, "Mapping");

var mapping = ds.Tables["Mapping"].AsEnumerable().Select(x => new {
    oldLogin = x.Field<string>(migrateUsersOldLoginColumn),
    newLogin = x.Field<string>(migrateUsersNewLoginColumn),
});

adapter.Dispose();


// sql scripts

var sqlCreateTempMappingTable = "CREATE TABLE #MappingTable (OldLogin NVARCHAR(256), NewLogin NVARCHAR(256))";

var mappingToSqlValues = mapping.Select(m => $"('{ m.oldLogin }', '{ m.newLogin }')");
var sqlInsertTempMappingTable = $"INSERT INTO #MappingTable VALUES { string.Join(", ", mappingToSqlValues) }";

var sqlUpdateUsersTable =   $"UPDATE { databaseUsersTable } " +
                            $"SET { databaseUsersLoginColumn } = " +
                                $"(SELECT TOP(1) NewLogin " +
                                $"FROM #MappingTable " +
                                $"WHERE { databaseUsersTable }.{ databaseUsersLoginColumn } = #MappingTable.OldLogin) " +
                            $"WHERE { databaseUsersLoginColumn } IN " +
                                $"(SELECT OldLogin " +
                                $"FROM #MappingTable)";

var sqlDropTempMappingTable = "DROP TABLE #MappingTable";


// update db users table

using (SqlConnection connection = new SqlConnection(databaseConnectionString))
{
    await connection.OpenAsync();

    SqlCommand command = new SqlCommand();
    command.Connection = connection;

    command.CommandText = sqlCreateTempMappingTable;
    await command.ExecuteNonQueryAsync();

    command.CommandText = sqlInsertTempMappingTable;
    await command.ExecuteNonQueryAsync();

    command.CommandText = sqlUpdateUsersTable;
    int number = await command.ExecuteNonQueryAsync();
    Console.WriteLine($"({ number } row affected)");

    command.CommandText = sqlDropTempMappingTable;
    await command.ExecuteNonQueryAsync();
}


Console.WriteLine("\n\nPress any button...");
Console.ReadKey();