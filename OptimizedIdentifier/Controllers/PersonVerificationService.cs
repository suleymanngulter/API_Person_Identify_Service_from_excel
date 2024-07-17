using ClosedXML.Excel;
using Newtonsoft.Json;
using OptimizedIdentifier.Model;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using verifyTc.Services;

namespace OptimizedIdentifier.Services
{
    public interface IPersonVerificationService
    {
        Task<string> VerifyFromExcelAsync(string filePath);
    }

    public class PersonVerificationService : IPersonVerificationService
    {
        private readonly IVerifyTc _verifyTc;

        public PersonVerificationService(IVerifyTc verifyTc)
        {
            _verifyTc = verifyTc;
        }

        public async Task<string> VerifyFromExcelAsync(string filePath)
        {
            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheets.First();
                var people = worksheet.RowsUsed().Skip(1).Select(row => new Person
                {
                    TcKimlikNo = long.Parse(row.Cell(1).Value.ToString()),
                    Name = row.Cell(2).Value.ToString(),
                    Surname = row.Cell(3).Value.ToString(),
                    YearOfBirth = int.Parse(row.Cell(4).Value.ToString())
                }).ToList();

                var verificationResults = await _verifyTc.Check(people);

                var allMessages = people.Select((person, index) => new
                {
                    PersonId = person.TcKimlikNo,
                    Name = person.Name,
                    Surname = person.Surname,
                    YearOfBirth = person.YearOfBirth,
                    IsVerified = verificationResults[index]
                }).Where(m => !m.IsVerified).ToList();

                return JsonConvert.SerializeObject(allMessages);
            }
        }
    }
}
