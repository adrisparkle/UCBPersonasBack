using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using UcbBack.Models;

namespace UcbBack.Logic
{
    public class ValidatePerson
    {
        private ApplicationDbContext _context;

        public ValidatePerson(ApplicationDbContext context)
        {
            _context = context;
        }
        public People UcbCode(People person)
        {
            int haszeroday = person.BIRTHDATE.ToString()[1] == '/' ? 1 : 0;
            int haszeromonth = person.BIRTHDATE.ToString()[4 - haszeroday] == '/' ? 1 : 0;
            char[] letras =
            {
                person.FIRSTSURNAME[0], 
                person.SECONDSURNAME[0], 
                person.NAMES[0], 
                '-',
                //day
                haszeroday==1? '0': person.BIRTHDATE.ToString()[0], 
                person.BIRTHDATE.ToString()[1-haszeroday],
                //month
                haszeromonth==1? '0': person.BIRTHDATE.ToString()[3-haszeroday], 
                person.BIRTHDATE.ToString()[4-haszeroday-haszeromonth], 
                //year
                person.BIRTHDATE.ToString()[8-haszeroday-haszeromonth], 
                person.BIRTHDATE.ToString()[9-haszeroday-haszeromonth]
                
                
            };

            person.COD_UCB = new string(letras);
            //colision!
            if ((_context.Person.FirstOrDefault(p => p.COD_UCB == person.COD_UCB))!= null)
            {
                char[] month =
                {
                    person.COD_UCB[6], person.COD_UCB[7]
                };
                int newmonth = Int32.Parse(new string(month)) > 12 ? Int32.Parse(new string(month)) + 10 : Int32.Parse(new string(month)) + 20;
                char[] oldCodArray = person.COD_UCB.ToCharArray();
                oldCodArray[6] = newmonth.ToString()[0];
                oldCodArray[7] = newmonth.ToString()[1];
                person.COD_UCB = new string(oldCodArray);
            }
            return person;
        }
    }
}