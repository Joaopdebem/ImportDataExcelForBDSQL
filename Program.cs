using System;
using Microsoft.Data.SqlClient;
using OfficeOpenXml;

class Program
{
    static void Main()
    {
        string connectionString = @"Server=localhost,1433;
                                    Database=bdManga;
                                    User ID=sa;
                                    Password=********;
                                    Trusted_Connection=False; 
                                    TrustServerCertificate=True;";

        string excelFilePath = "./Mangas.xlsx";

        using (var package = new ExcelPackage(new System.IO.FileInfo(excelFilePath)))
        {
            var worksheet = package.Workbook.Worksheets[0];

            int rowCount = worksheet.Dimension.Rows;

            using (var connection = new SqlConnection(connectionString))
            {
                connection.Open();

                for (int row = 2; row <= rowCount; row++)
                {
                    int categoryId = string.IsNullOrWhiteSpace(worksheet.Cells[row, 1].Text) ? 0 : Convert.ToInt32(worksheet.Cells[row, 1].Text);
                    int typeId = string.IsNullOrWhiteSpace(worksheet.Cells[row, 2].Text) ? 0 : Convert.ToInt32(worksheet.Cells[row, 1].Text);
                    string title = worksheet.Cells[row, 3].Text;
                    string description = worksheet.Cells[row, 4].Text;
                    decimal chapter = Convert.ToDecimal(worksheet.Cells[row, 5].Text);
                    int rating = Convert.ToInt32(worksheet.Cells[row, 6].Text);
                    char isFinished = Convert.ToChar(worksheet.Cells[row, 7].Text);
                    char isFavorite = Convert.ToChar(worksheet.Cells[row, 8].Text);
                    string slug = worksheet.Cells[row, 9].Text;

                    string insertCommand = "INSERT INTO Collection (CategoryId, TypeId, Title, Description, Chapter, Rating, isFinished, isFavorite, Slug) " +
                                           "VALUES (@CategoryId, @TypeId, @Title, @Description, @Chapter, @Rating, @IsFinished, @IsFavorite, @Slug)";

                    using (var sqlCommand = new SqlCommand(insertCommand, connection))
                    {
                        sqlCommand.Parameters.AddWithValue("@CategoryId", categoryId);
                        sqlCommand.Parameters.AddWithValue("@TypeId", typeId);
                        sqlCommand.Parameters.AddWithValue("@Title", title);
                        sqlCommand.Parameters.AddWithValue("@Description", description);
                        sqlCommand.Parameters.AddWithValue("@Chapter", chapter);
                        sqlCommand.Parameters.AddWithValue("@Rating", rating);
                        sqlCommand.Parameters.AddWithValue("@IsFinished", isFinished);
                        sqlCommand.Parameters.AddWithValue("@IsFavorite", isFavorite);
                        sqlCommand.Parameters.AddWithValue("@Slug", slug);

                        sqlCommand.ExecuteNonQuery();
                    }
                }

                connection.Close();
            }
        }

        Console.WriteLine("Importação concluída com sucesso!");
    }
}
