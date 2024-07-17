using KPSPublic;
using OptimizedIdentifier.Model;

namespace verifyTc.Services
{
    public interface IVerifyTc
    {
        Task<bool[]> Check(List<Person> people);

    }

    public class VerifyTc : IVerifyTc
    {
        public async Task<bool[]> Check(List<Person> people)
        {
            var service = new KPSPublicSoapClient(KPSPublicSoapClient.EndpointConfiguration.KPSPublicSoap);
            var results = new bool[people.Count];

            for (int i = 0; i < people.Count; i++)
            {
                var person = people[i];
                var response = await service.TCKimlikNoDogrulaAsync(person.TcKimlikNo, person.Name, person.Surname, person.YearOfBirth);
                results[i] = response.Body.TCKimlikNoDogrulaResult;
            }

            return results;
        }
    }
}