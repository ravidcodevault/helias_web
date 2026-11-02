using System.Data.SqlClient;

namespace OYA_Web_Application.Models.Services
{
    public class DistributorService
    {
        private readonly string _connectionString;
        private readonly IConfiguration _configuration;

        public DistributorService(IConfiguration configuration)
        {
            _configuration = configuration;
            _connectionString =  _configuration.GetConnectionString("DefaultConnection");
        }
        public List<DistributorModel> GetAll()
        {
            var list = new List<DistributorModel>();

            using (SqlConnection conn = new SqlConnection(_connectionString))
            {
                conn.Open();

                using (SqlCommand cmd = new SqlCommand("SELECT * FROM Distributor", conn))
                using (SqlDataReader reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        list.Add(new DistributorModel
                        {
                            Customer = reader["Customer"]?.ToString(),
                            Name = reader["name"]?.ToString(),
                            Code = reader["code"]?.ToString(),
                            Sheet = reader["Sheet"]?.ToString(),
                            Region = reader["region"]?.ToString(),
                            Territory = reader["territory"]?.ToString(),
                            HQ = reader["HQ"]?.ToString(),
                            Row = reader["Row"] != DBNull.Value ? Convert.ToInt32(reader["Row"]) : 0,
                            WestRow = reader["WestRow"] != DBNull.Value ? Convert.ToInt32(reader["WestRow"]) : 0,
                            EastRow = reader["EastRow"] != DBNull.Value ? Convert.ToInt32(reader["EastRow"]) : 0,
                            CentralRow = reader["CentralRow"] != DBNull.Value ? Convert.ToInt32(reader["CentralRow"]) : 0,
                            InternationalRow = reader["InternationalRow"] != DBNull.Value ? Convert.ToInt32(reader["InternationalRow"]) : 0
                        });
                    }
                }
            }

            return list;
        }
        public void Insert(DistributorModel model)
        {
            using (SqlConnection conn = new SqlConnection(_connectionString))
            {
                conn.Open();

                var query = @"
        INSERT INTO Distributor
        (Customer, name, code, Sheet, region, territory, HQ, Row, WestRow, EastRow, CentralRow, InternationalRow, PrevYearTotal, Plan)
        VALUES
        (@Customer, @Name, @Code, @Sheet, @Region, @Territory, @HQ, @Row, @WestRow, @EastRow, @CentralRow, @InternationalRow, @PrevYearTotal, @Plan)";

                using (SqlCommand cmd = new SqlCommand(query, conn))
                {
                    cmd.Parameters.AddWithValue("@Customer", model.Customer ?? "");
                    cmd.Parameters.AddWithValue("@Name", model.Name ?? "");
                    cmd.Parameters.AddWithValue("@Code", model.Code ?? "");
                    cmd.Parameters.AddWithValue("@Sheet", model.Sheet ?? "");
                    cmd.Parameters.AddWithValue("@Region", model.Region ?? "");
                    cmd.Parameters.AddWithValue("@Territory", model.Territory ?? "");
                    cmd.Parameters.AddWithValue("@HQ", model.HQ ?? "");
                    cmd.Parameters.AddWithValue("@Row", model.Row);
                    cmd.Parameters.AddWithValue("@WestRow", model.WestRow);
                    cmd.Parameters.AddWithValue("@EastRow", model.EastRow);
                    cmd.Parameters.AddWithValue("@CentralRow", model.CentralRow);
                    cmd.Parameters.AddWithValue("@InternationalRow", model.InternationalRow);
                    cmd.Parameters.AddWithValue("@PrevYearTotal", model.PrevYearTotal ?? (object)DBNull.Value);
                    cmd.Parameters.AddWithValue("@Plan", model.Plan ?? (object)DBNull.Value);

                    cmd.ExecuteNonQuery();
                }
            }
        }
        public void Update(DistributorModel model)
        {
            using (SqlConnection conn = new SqlConnection(_connectionString))
            {
                conn.Open();

                var query = @"
        UPDATE Distributor SET
            name = @Name,
            code = @Code,
            Sheet = @Sheet,
            region = @Region,
            territory = @Territory,
            HQ = @HQ,
            Row = @Row,
            WestRow = @WestRow,
            EastRow = @EastRow,
            CentralRow = @CentralRow,
            InternationalRow = @InternationalRow,
            PrevYearTotal = @PrevYearTotal,
            Plan = @Plan
        WHERE Customer = @Customer";

                using (SqlCommand cmd = new SqlCommand(query, conn))
                {
                    cmd.Parameters.AddWithValue("@Customer", model.Customer);
                    cmd.Parameters.AddWithValue("@Name", model.Name ?? "");
                    cmd.Parameters.AddWithValue("@Code", model.Code ?? "");
                    cmd.Parameters.AddWithValue("@Sheet", model.Sheet ?? "");
                    cmd.Parameters.AddWithValue("@Region", model.Region ?? "");
                    cmd.Parameters.AddWithValue("@Territory", model.Territory ?? "");
                    cmd.Parameters.AddWithValue("@HQ", model.HQ ?? "");
                    cmd.Parameters.AddWithValue("@Row", model.Row);
                    cmd.Parameters.AddWithValue("@WestRow", model.WestRow);
                    cmd.Parameters.AddWithValue("@EastRow", model.EastRow);
                    cmd.Parameters.AddWithValue("@CentralRow", model.CentralRow);
                    cmd.Parameters.AddWithValue("@InternationalRow", model.InternationalRow);
                    cmd.Parameters.AddWithValue("@PrevYearTotal", model.PrevYearTotal ?? (object)DBNull.Value);
                    cmd.Parameters.AddWithValue("@Plan", model.Plan ?? (object)DBNull.Value);

                    cmd.ExecuteNonQuery();
                }
            }
        }
        public List<DistributorModel> GetAllMaster()
        {
            var list = new List<DistributorModel>();

            using (SqlConnection conn = new SqlConnection(_connectionString))
            {
                conn.Open();

                var query = @"
            SELECT Customer, name, code, region, territory
            FROM Distributor";

                using (SqlCommand cmd = new SqlCommand(query, conn))
                using (SqlDataReader reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        list.Add(new DistributorModel
                        {
                            Customer = reader["Customer"]?.ToString(),
                            Name = reader["name"]?.ToString(),
                            Code = reader["code"]?.ToString(),
                            Region = reader["region"]?.ToString(),
                            Territory = reader["territory"]?.ToString()
                        });
                    }
                }
            }

            return list;
        }
        public void InsertMaster(DistributorModel model)
        {
            using (SqlConnection conn = new SqlConnection(_connectionString))
            {
                conn.Open();

                var query = @"
            INSERT INTO Distributor
            (Customer, name, code, region, territory)
            VALUES
            (@Customer, @Name, @Code, @Region, @Territory)";

                using (SqlCommand cmd = new SqlCommand(query, conn))
                {
                    cmd.Parameters.AddWithValue("@Customer", model.Customer ?? "");
                    cmd.Parameters.AddWithValue("@Name", model.Name ?? "");
                    cmd.Parameters.AddWithValue("@Code", model.Code ?? "");
                    cmd.Parameters.AddWithValue("@Region", model.Region ?? "");
                    cmd.Parameters.AddWithValue("@Territory", model.Territory ?? "");

                    cmd.ExecuteNonQuery();
                }
            }
        }
        public void UpdateMaster(DistributorModel model)
        {
            using (SqlConnection conn = new SqlConnection(_connectionString))
            {
                conn.Open();

                var query = @"
            UPDATE Distributor SET
                name = @Name,
                code = @Code,
                region = @Region,
                territory = @Territory
            WHERE Customer = @Customer";

                using (SqlCommand cmd = new SqlCommand(query, conn))
                {
                    cmd.Parameters.AddWithValue("@Customer", model.Customer);
                    cmd.Parameters.AddWithValue("@Name", model.Name ?? "");
                    cmd.Parameters.AddWithValue("@Code", model.Code ?? "");
                    cmd.Parameters.AddWithValue("@Region", model.Region ?? "");
                    cmd.Parameters.AddWithValue("@Territory", model.Territory ?? "");

                    cmd.ExecuteNonQuery();
                }
            }
        }
        public void ApplyAutoShift(string type, int newRow)
        {
            if (newRow <= 0) return;

            using (SqlConnection conn = new SqlConnection(_connectionString))
            {
                conn.Open();

                string column = type switch
                {
                    "main" => "Row",
                    "west" => "WestRow",
                    "east" => "EastRow",
                    "central" => "CentralRow",
                    "intl" => "InternationalRow",
                    _ => ""
                };

                if (string.IsNullOrEmpty(column)) return;

                var query = $@"
            UPDATE Distributor
            SET {column} = {column} + 1
            WHERE {column} >= @Row";

                using (SqlCommand cmd = new SqlCommand(query, conn))
                {
                    cmd.Parameters.AddWithValue("@Row", newRow);
                    cmd.ExecuteNonQuery();
                }
            }
        }
        public void UpdateMappingWithShift(List<DistributorModel> list)
        {
            using (SqlConnection conn = new SqlConnection(_connectionString))
            {
                conn.Open();

                foreach (var d in list)
                {
                    if (d._changed && d._newRow > 0)
                    {
                        string column = d._type switch
                        {
                            "main" => "Row",
                            "west" => "WestRow",
                            "east" => "EastRow",
                            "central" => "CentralRow",
                            "intl" => "InternationalRow",
                            _ => ""
                        };

                        if (!string.IsNullOrEmpty(column))
                        {
                            var shiftQuery = $@"
                        UPDATE Distributor
                        SET {column} = {column} + 1
                        WHERE {column} >= @Row";

                            using (SqlCommand shiftCmd = new SqlCommand(shiftQuery, conn))
                            {
                                shiftCmd.Parameters.AddWithValue("@Row", d._newRow);
                                shiftCmd.ExecuteNonQuery();
                            }
                        }
                    }

                    var updateQuery = @"
                UPDATE Distributor SET
                    Row = @Row,
                    WestRow = @WestRow,
                    EastRow = @EastRow,
                    CentralRow = @CentralRow,
                    InternationalRow = @InternationalRow
                WHERE Customer = @Customer";

                    using (SqlCommand cmd = new SqlCommand(updateQuery, conn))
                    {
                        cmd.Parameters.AddWithValue("@Customer", d.Customer ?? "");
                        cmd.Parameters.AddWithValue("@Row", d.Row);
                        cmd.Parameters.AddWithValue("@WestRow", d.WestRow);
                        cmd.Parameters.AddWithValue("@EastRow", d.EastRow);
                        cmd.Parameters.AddWithValue("@CentralRow", d.CentralRow);
                        cmd.Parameters.AddWithValue("@InternationalRow", d.InternationalRow);

                        cmd.ExecuteNonQuery();
                    }
                }
            }
        }
        public void UpdateMapping(List<DistributorModel> list)
        {
            using (SqlConnection conn = new SqlConnection(_connectionString))
            {
                conn.Open();

                foreach (var d in list)
                {
                    var query = @"
                UPDATE Distributor SET
                    Row = @Row,
                    WestRow = @WestRow,
                    EastRow = @EastRow,
                    CentralRow = @CentralRow,
                    InternationalRow = @InternationalRow
                WHERE Customer = @Customer";

                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    {
                        cmd.Parameters.AddWithValue("@Customer", d.Customer ?? "");
                        cmd.Parameters.AddWithValue("@Row", d.Row);
                        cmd.Parameters.AddWithValue("@WestRow", d.WestRow);
                        cmd.Parameters.AddWithValue("@EastRow", d.EastRow);
                        cmd.Parameters.AddWithValue("@CentralRow", d.CentralRow);
                        cmd.Parameters.AddWithValue("@InternationalRow", d.InternationalRow);

                        cmd.ExecuteNonQuery();
                    }
                }
            }
        }
        public void AddMappingWithShift(MappingRequest req)
        {
            using var conn = new SqlConnection(_connectionString);
            conn.Open();

            string column = req.Type switch
            {
                "main" => "Row",
                "west" => "WestRow",
                "east" => "EastRow",
                "central" => "CentralRow",
                "intl" => "InternationalRow",
                _ => ""
            };

            if (string.IsNullOrEmpty(column)) return;

            // SHIFT
            var shift = $@"UPDATE Distributor SET {column} = {column} + 1 WHERE {column} >= @Row";
            using (var cmd = new SqlCommand(shift, conn))
            {
                cmd.Parameters.AddWithValue("@Row", req.Row);
                cmd.ExecuteNonQuery();
            }

            // UPDATE TARGET
            var update = $@"UPDATE Distributor SET {column} = @Row WHERE Customer = @Customer";
            using (var cmd = new SqlCommand(update, conn))
            {
                cmd.Parameters.AddWithValue("@Customer", req.Customer);
                cmd.Parameters.AddWithValue("@Row", req.Row);
                cmd.ExecuteNonQuery();
            }
        }
    }
}
