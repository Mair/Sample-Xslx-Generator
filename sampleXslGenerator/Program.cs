using GeneratedCode;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace sampleXslGenerator
{
    class Program
    {
        static void Main(string[] args)
        {
            new GeneratedClass().CreatePackage("mySpreadSheet.xlsx");
            var PopulateSpreadSheet = new PopulateSpreadSheet(ScafaldPerson());
            using (var stream = File.Open("mySpreadSheet.xlsx", FileMode.Open))
            {
                PopulateSpreadSheet.Polpulate(stream);
            }
        }

        static List<Person> ScafaldPerson()
        {
            var people = new List<Person>();
            for (int i = 0; i < 100000; i++)
            {
                people.Add(new Person(){ 
                    FirstName="FirstName " + i, 
                    LastName="Lastname " + i,
                    DateOfBirth = new DateTime(2007,07,20).AddYears(i % 10)});
            }
            return people;
        }
    }
}
