using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Excel_Demo_Eppuls_Self.Models;
using System.Linq;

namespace Excel_Demo_Eppuls_Self
{
    class Program
    {
        static async Task Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var file = new FileInfo(@"..\demo.xlsx");

            var users = GetUsersList();
            try
            {
                await CreateExcelFile(file, users);
                Console.WriteLine("File has been saved.");
            }
            catch
            {
                Console.WriteLine($"The file{file.Name} can not be saved. Check the file is open.");
            }

            try
            {
                var usersfromExcel = await ReadDataFromExcel(file);

                foreach(User user in usersfromExcel)
                {
                    Console.WriteLine($"{user.Id} {user.FirstName} {user.LastName} {user.Position}");
                }
                

            }
            catch
            {
                Console.WriteLine($"File  {file.Name} can not be opened.");
            }
            


        }

        private static async Task<List<User>> ReadDataFromExcel(FileInfo file)
        {
            using var package = new ExcelPackage(file);

            int row = 2;
            int col = 1;

            List<User> usersFromExcel = new();

            await package.LoadAsync(file);

            var ws = package.Workbook.Worksheets[0];

            while(string.IsNullOrWhiteSpace(ws.Cells[row, col].Value?.ToString()) == false)
            {
                User u = new();
                u.Id = int.Parse(ws.Cells[row, col].Value.ToString());
                u.FirstName = ws.Cells[row, col + 1].Value.ToString();
                u.LastName = ws.Cells[row, col + 2].Value.ToString();
                u.Position = ws.Cells[row, col + 3].Value.ToString();

                usersFromExcel.Add(u);
                row += 1;
            }

            return usersFromExcel;
        }

        private static async Task CreateExcelFile(FileInfo file, List<User> users)
        {

            DeleteFile(file);

            using var package = new ExcelPackage(file);
            var ws = package.Workbook.Worksheets.Add("Arkusz1");
            ws.Cells["A1"].LoadFromCollection(users,true);
            ws.Row(1).Style.Font.Bold = true;

            await package.SaveAsync();
            
        }

        private static void DeleteFile(FileInfo file)
        {
            if(file.Exists)
            {
                file.Delete();
            }
        }

        private static List<User> GetUsersList()
        {
            List<User> usres = new()
            {
                new User { Id = 1, FirstName ="Łukasz",LastName="Kapiński",Position="Specialista"},
                new User { Id = 2, FirstName = "Jan", LastName = "Sapiński", Position = "Specialista" },
                new User { Id = 3, FirstName = "Mateusz", LastName = "Nowak", Position = "Specialista" },
                new User { Id = 4, FirstName = "Roman", LastName = "Kowalski", Position = "Koordynator" }
            };

            return usres;
        }
    }
}
