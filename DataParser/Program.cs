using Newtonsoft.Json.Linq;
using System;
using System.IO;
using System.Threading;

namespace DataParser
{
    class Program
    {
        static void Main(string[] args)
        {
            string json = File.ReadAllText("people-0.json");
            dynamic people = JArray.Parse(json);
            int id = 0;

            using StreamWriter sw = new StreamWriter("people-00.json");

            sw.WriteLine("[");
            foreach(var person in people)
            {
                id++;
                string[] names = person.name.ToString().Split(" ");
                person.id = id;
                person.lastname = names[0];
                person.name = names[1];
                person.father = names[2];

                sw.WriteLine(person.ToString()+",");
            }
            sw.WriteLine("]");


        }
    }
}
