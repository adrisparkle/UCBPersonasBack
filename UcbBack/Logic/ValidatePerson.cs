using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Web;
using UcbBack.Models;

namespace UcbBack.Logic
{
    public class ValidatePerson
    {
        private ApplicationDbContext _context;
        private HanaValidator hanaValidator;

        public ValidatePerson(ApplicationDbContext context)
        {
            _context = context;
            hanaValidator = new HanaValidator(_context);
        }

        public People CleanName(People person)
        {
            person.FirstSurName = hanaValidator.CleanText(person.FirstSurName);
            person.SecondSurName = hanaValidator.CleanText(person.SecondSurName);
            person.Names = hanaValidator.CleanText(person.Names);
            person.MariedSurName = hanaValidator.CleanText(person.MariedSurName);

            return person;
        }

        

        public string GetConfirmationToken(IEnumerable<People> people)
        {
            string token = "";

            foreach (var p in people)
            {
                token += p.CUNI;
            }

            byte[] encodedstr = new UTF8Encoding().GetBytes(token);
            byte[] hash = ((HashAlgorithm)CryptoConfig.CreateFromName("MD5")).ComputeHash(encodedstr);
            token = Convert.ToBase64String(hash);

            return token;
        }

        public People UcbCode(People person)
        {
            DateTime bd = person.BirthDate;
            var daystr = person.BirthDate.ToString("dd");
            var monthstr = person.BirthDate.ToString("MM");
            var yearstr = person.BirthDate.ToString("yy");

            char[] letras =
            {
                person.FirstSurName[0], 
                person.SecondSurName[0], 
                person.Names[0], 
                //#'-',
                //year
                yearstr[0],
                yearstr[1],
                //month
                monthstr[0],
                monthstr[1],
                //day
                daystr[0],
                daystr[1]
            };

            person.CUNI = new string(letras);
            //colision!
            if ((_context.Person.FirstOrDefault(p => p.CUNI == person.CUNI)) != null)
            {
                char[] monthi =
                {
                    person.CUNI[5], person.CUNI[6]
                };
                int newmonth = Int32.Parse(new string(monthi)) > 12 ? Int32.Parse(new string(monthi)) + 10 : Int32.Parse(new string(monthi)) + 20;
                char[] oldCodArray = person.CUNI.ToCharArray();
                oldCodArray[5] = newmonth.ToString()[0];
                oldCodArray[6] = newmonth.ToString()[1];
                person.CUNI = new string(oldCodArray);
            }
            return person;
        }

        public IEnumerable<People> VerifyExisting(People person, float n,string fn=null)
        {
            string fullname;
            if (fn == null)
            {
                fullname = String.Concat(person.FirstSurName,
                    String.Concat(" ",
                        String.Concat(person.SecondSurName,
                            String.Concat(" ",
                                String.Concat(person.Names,
                                    String.Concat(" ", person.Document)
                                )
                            )
                        )
                    )
                );
            }
            else
                fullname = fn;
            
            //SQL command in Hana
            string colToCompare = "concat(a.\"FirstSurName\"," +
                                "concat('' '',"+
                                    "concat(a.\"SecondSurName\","+
                                        "concat('' '',"+
                                            "concat(a.\"Names\", "+
                                                "concat('' '',a.\"Document\")"+
                                            ")"+
                                        ")"+
                                    ")"+
                                ")"+
                            ")";
            string colId = "a.\"CUNI\"";
            string table = "People";
            
            var similarities = hanaValidator.Similarities(fullname, colToCompare, table, colId, 0.9f);

            var sim = _context.Person.Where(
                i => similarities.Contains(i.CUNI)
            ).ToList();
            return sim;
        }
    }
}