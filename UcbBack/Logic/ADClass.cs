using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.Linq;
using System.Web;
using UcbBack.Models;
using UcbBack.Models.Auth;
using System.Data.Entity;

namespace UcbBack.Logic
{
    public class ADClass
    {
        // first install directory services dll
        //Install-Package System.DirectoryServices -Version 4.5.0
        //Install-Package System.DirectoryServices.AccountManagement -Version 4.5.0

        public string sDomain = "UCB.BO";
        public string Domain = "192.168.18.62";
        public void addUser(People person)
        {
            try
            {
                var contract = person.GetLastContract();

                var branchGroup = contract.Branches.ADGroupName;

                using (PrincipalContext ouContex = new PrincipalContext(ContextType.Domain,
                    Domain,
                    "OU=" + contract.Branches.ADOUName + ",DC=UCB,DC=BO",
                    "ADMNALRRHH",
                    "Rrhh1234"))
                {
                    //ouContex.ValidateCredentials("ADMNALRRHH", "Rrhh1234");
                    using (UserPrincipal up = new UserPrincipal(ouContex))
                    {
                        var initials = person.FirstSurName.ToCharArray()[0].ToString() +
                                       person.Names.ToCharArray()[0].ToString();
                        up.GivenName = person.Names;
                        //up.MiddleName = person.SecondSurName;
                        up.Surname = person.FirstSurName;
                        up.DisplayName = person.GetFullName();
                        up.Name = up.DisplayName;
                        up.SamAccountName = getSamAcoutName(person);
                        up.UserPrincipalName = up.SamAccountName + "@UCB.BO";
                        up.SetPassword(person.Document); // user ChangePassword to change password lol
                        up.VoiceTelephoneNumber = person.PhoneNumber;
                        up.EmailAddress = person.UcbEmail;
                        up.EmployeeId = person.CUNI;
                        up.Enabled = true;
                        //up.ExpirePasswordNow();


                        up.Save();
                        AddUserToGroup(up.UserPrincipalName, contract.Branches.ADGroupName); // allways with UserPrincipalName to add to the group

                        if (up.GetUnderlyingObjectType() == typeof(DirectoryEntry))
                        {
                            using (var entry = (DirectoryEntry)up.GetUnderlyingObject())
                            {
                                entry.Properties["initials"].Value = initials;
                                entry.Properties["title"].Value = contract.Positions.Name;
                                entry.Properties["company"].Value = contract.Branches.Name;
                                //todo find a way to know who is the manager
                                //entry.Properties["manager"].Value = "NaN";
                                entry.Properties["department"].Value = contract.Dependency.Name;
                                entry.CommitChanges();
                            }
                        }
                    }
                }
            }
            catch (PrincipalExistsException e)
            {
                Console.WriteLine(e);
            }
            
        }

        public string getSamAcoutName(People person)
        {
            var _context = new ApplicationDbContext();
            var personuser = _context.CustomUsers.FirstOrDefault(x => x.PeopleId == person.Id);


            if (personuser!=null)
            {
                return personuser.UserPrincipalName.Split('@')[0];
            }
            // First attempt
            var SAN = person.Names.ToCharArray()[0].ToString() + "."
                      + person.FirstSurName
                      + (person.SecondSurName != null ? ("." + person.SecondSurName.ToCharArray()[0].ToString()) : "");
            SAN = System.Text.Encoding.UTF8.GetString(System.Text.Encoding.GetEncoding("ISO-8859-8").GetBytes(SAN.Replace(" ", "")));
            var UPN = SAN + "@" + Domain;
            var search = _context.CustomUsers.Where(x => x.UserPrincipalName == UPN).ToList();
            if (search.Any())
            {
                // Second attempt
                SAN = person.Names.ToCharArray()[0].ToString() + person.Names.ToCharArray()[1].ToString() + "."
                        + person.FirstSurName
                        + (person.SecondSurName != null ? ("." + person.SecondSurName.ToCharArray()[0].ToString()) : "");
                SAN = System.Text.Encoding.UTF8.GetString(System.Text.Encoding.GetEncoding("ISO-8859-8").GetBytes(SAN.Replace(" ", "")));

                UPN = SAN + "@" + Domain;
                search = _context.CustomUsers.Where(x => x.UserPrincipalName == UPN).ToList();
                if (search.Any())
                {
                    // Third attempt
                    SAN = person.Names.ToCharArray()[0].ToString() + person.Names.ToCharArray()[1].ToString() + "."
                          + person.FirstSurName
                          + (person.SecondSurName != null ? ("." + person.SecondSurName.ToCharArray()[0].ToString() + person.SecondSurName.ToCharArray()[1].ToString()) : "");
                    SAN = System.Text.Encoding.UTF8.GetString(System.Text.Encoding.GetEncoding("ISO-8859-8").GetBytes(SAN.Replace(" ", "")));

                    UPN = SAN + "@" + Domain;
                    search = _context.CustomUsers.Where(x => x.UserPrincipalName == UPN).ToList();
                    if (search.Any())
                    {
                        // Fourth attempt
                        SAN = person.Names.ToCharArray()[0].ToString() + "."
                              + person.FirstSurName
                              + (person.SecondSurName != null ? ("." + person.SecondSurName.ToCharArray()[0].ToString()) : "")
                              + (person.BirthDate.Day).ToString().PadLeft(2,'0');
                        SAN = System.Text.Encoding.UTF8.GetString(System.Text.Encoding.GetEncoding("ISO-8859-8").GetBytes(SAN.Replace(" ", "")));

                        UPN = SAN + "@" + Domain;
                        Random rnd = new Random();
                        while (_context.CustomUsers.Where(x => x.UserPrincipalName == UPN).ToList().Any())
                        {
                            // Last and final attempt

                            SAN = person.Names.ToCharArray()[0].ToString() + "."
                                                                + person.FirstSurName
                                                                + (person.SecondSurName != null ? ("." + person.SecondSurName.ToCharArray()[0].ToString()) : "")
                                                                + (rnd.Next(1, 100)).ToString().PadLeft(2, '0');
                            SAN = System.Text.Encoding.UTF8.GetString(System.Text.Encoding.GetEncoding("ISO-8859-8").GetBytes(SAN.Replace(" ", "")));
                        }
                    }
                }
            }

            //person.UserPrincipalName = UPN;

            var newUser = new CustomUser();
            newUser.PeopleId = person.Id;
            newUser.UserPrincipalName = UPN;

            return SAN;
        }

        public bool memberOf(CustomUser user,string groupName)
        {
            using (PrincipalContext ouContex = new PrincipalContext(ContextType.Domain,
                Domain,
                "ADMNALRRHH@UCB.BO",
                "Rrhh1234"))
            {
                ouContex.ValidateCredentials("ADMNALRRHH", "Rrhh1234");
                var u = findUser(user.People);
                GroupPrincipal group = GroupPrincipal.FindByIdentity(ouContex, groupName);
                if (group != null)
                {
                    return (u.IsMemberOf(group));
                }
            }

            return false;
        }

        public Principal findUser(People person)
        {
            Principal user;
            PrincipalContext ouContex = new PrincipalContext(ContextType.Domain,
                Domain,
                "ADMNALRRHH@UCB.BO",
                "Rrhh1234");
            ouContex.ValidateCredentials("ADMNALRRHH", "Rrhh1234");
            UserPrincipal up = new UserPrincipal(ouContex);    
            up.EmployeeId = person.CUNI;
            PrincipalSearcher ps = new PrincipalSearcher(up);
            user = (UserPrincipal)ps.FindOne();
            return user;
        }

        public List<Branches> getUserBranches(CustomUser customUser)
        {
            List<Branches> roles = new List<Branches>();
            PrincipalContext ouContex = new PrincipalContext(ContextType.Domain,
                Domain,
                "ADMNALRRHH@UCB.BO",
                "Rrhh1234");

            ouContex.ValidateCredentials("ADMNALRRHH", "Rrhh1234");
            UserPrincipal user = UserPrincipal.FindByIdentity(ouContex, customUser.UserPrincipalName);
            if (user != null)
            {
                List<string> grps = new List<string>();
                var groups = user.GetGroups();
                foreach (var group in groups)
                {
                    grps.Add(group.Name);
                }
                var _context = new ApplicationDbContext();
                roles = _context.Branch.ToList().Where(x => grps.Contains(x.ADGroupName)).ToList();
            }

            return roles;
        }

        public List<Rol> getUserRols(CustomUser customUser)
        {
            List<Rol> roles = new List<Rol>();
            PrincipalContext ouContex = new PrincipalContext(ContextType.Domain,
                Domain,
                "ADMNALRRHH@UCB.BO",
                "Rrhh1234");
            
            ouContex.ValidateCredentials("ADMNALRRHH", "Rrhh1234");
            UserPrincipal user = UserPrincipal.FindByIdentity(ouContex, customUser.UserPrincipalName);
            if (user != null)
            {
                List<string> grps = new List<string>();
                var groups = user.GetGroups();
                foreach (var group in groups)
                {
                    grps.Add(group.Name);
                }
                var _context = new ApplicationDbContext();
                roles = _context.Rols.Include(x=>x.Resource).ToList().Where(x => grps.Contains(x.ADGroupName)).ToList();
            }

            return roles;
        }

        public void AddUserToGroup(string userPrincipalName, string groupName)
        {
            try
            {
                using (PrincipalContext ouContex = new PrincipalContext(ContextType.Domain,
                    Domain,
                    "ADMNALRRHH@UCB.BO",
                    "Rrhh1234"))
                {
                    ouContex.ValidateCredentials("ADMNALRRHH", "Rrhh1234");
                    GroupPrincipal group = GroupPrincipal.FindByIdentity(ouContex, groupName);
                    group.Members.Add(ouContex, IdentityType.UserPrincipalName, userPrincipalName);
                    group.Save();
                }
            }
            catch (System.DirectoryServices.DirectoryServicesCOMException E)
            {
                //doSomething with E.Message.ToString(); 

            }
        }
        public void RemoveUserFromGroup(string userPrincipalName, string groupName)
        {
            try
            {
                using (PrincipalContext ouContex = new PrincipalContext(ContextType.Domain,
                    Domain,
                    "ADMNALRRHH@UCB.BO",
                    "Rrhh1234"))
                {
                    ouContex.ValidateCredentials("ADMNALRRHH", "Rrhh1234");
                    GroupPrincipal group = GroupPrincipal.FindByIdentity(ouContex, groupName);
                    group.Members.Remove(ouContex, IdentityType.UserPrincipalName, userPrincipalName);
                    group.Save();
                }
            }
            catch (System.DirectoryServices.DirectoryServicesCOMException E)
            {
                //doSomething with E.Message.ToString(); 

            }
        }

        public bool ActiveDirectoryAuthenticate(string username, string password)
        {

            PrincipalContext ouContex = new PrincipalContext(ContextType.Domain,
                Domain,
                "ADMNALRRHH@UCB.BO",
                "Rrhh1234");
                bool Valid = ouContex.ValidateCredentials(username, password);
                if (Valid)
                {
                    ouContex.ValidateCredentials("ADMNALRRHH", "Rrhh1234");
                    UserPrincipal up = new UserPrincipal(ouContex);
                        up.UserPrincipalName = username;
                        PrincipalSearcher ps = new PrincipalSearcher(up);
                        
                            var user = (UserPrincipal) ps.FindOne();
                            
                            return user == null ? false : user.Enabled.Value;
                        
                    }

                return false;
            
        }

        public List<string> getGroups()
        {
            List<string> res = new List<string>();
            PrincipalContext ouContex = new PrincipalContext(ContextType.Domain,
                Domain,
                "ADMNALRRHH@UCB.BO",
                "Rrhh1234");
            ouContex.ValidateCredentials("ADMNALRRHH", "Rrhh1234");
            GroupPrincipal qbeGroup = new GroupPrincipal(ouContex);
            
            PrincipalSearcher srch = new PrincipalSearcher(qbeGroup);

            

            foreach (var group in srch.FindAll())
            {
                res.Add(((GroupPrincipal)group).Name);   
            }

            return res;

        }

        public bool createGroup(string name)
        {
            using (PrincipalContext ouContex = new PrincipalContext(ContextType.Domain,
                Domain,
                "OU=Personas,DC=UCB,DC=BO",
                "ADMNALRRHH",
                "Rrhh1234"))
            {
                ouContex.ValidateCredentials("ADMNALRRHH", "Rrhh1234");
                GroupPrincipal group = GroupPrincipal.FindByIdentity(ouContex, name);
                if (group == null)
                {
                    using (GroupPrincipal up = new GroupPrincipal(ouContex))
                    {
                        up.IsSecurityGroup = false;
                        up.Name = name;
                        up.DisplayName = name;
                        up.GroupScope = GroupScope.Global;
                        up.Save();
                    }

                    return true;
                }

                return false;
            }
        }

        /*public bool ChangePassword(CustomUser customUser, string oldpassword, string newpassword)
        {
            UserPrincipal user = null;
            using (PrincipalContext ouContex = new PrincipalContext(ContextType.Domain,
                Domain,
                "ADMNALRRHH",
                "Rrhh1234"))
            {
                using (UserPrincipal up = new UserPrincipal(ouContex))
                {
                    up.EmployeeId = person.CUNI;
                    using (PrincipalSearcher ps = new PrincipalSearcher(up))
                    {
                        user = (UserPrincipal)ps.FindOne();
                    }
                }
            }

            if (user == null) return false;
            if (!ActiveDirectoryAuthenticate(person.UserPrincipalName, oldpassword)) return false;
            try
            {
                user.ChangePassword(oldpassword, newpassword);
                return true;
            }
            catch (Exception e)
            {
                return false;
            } 
        }*/


    }
}