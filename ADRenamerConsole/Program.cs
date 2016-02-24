using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.DirectoryServices;
using System.IO;

namespace ADRenamerConsole
{
    class Program
    {
        public static string LDAP_scope, profileFolder, userFolder;

        static void Main(string[] args)
        {
            #region Settings.ini check
            try
            {
                string[] Settings = File.ReadAllLines("settings.ini");
                LDAP_scope = Settings[0];
                profileFolder = Settings[1];
                userFolder = Settings[2];
                Console.CursorVisible = false;
            }
            catch
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Problem with settings.ini file, shutting down...");
                Environment.Exit(0);
            }
            #endregion

            Console.WriteLine("Retrieving users from: " + LDAP_scope + "...");

            var Users = new List<ADUserProperties>();

            using (DirectoryEntry de = new DirectoryEntry(LDAP_scope))
            {
                using (DirectorySearcher ds = new DirectorySearcher(de))
                {
                    ds.Filter = "(&(objectClass=user)(objectCategory=person))";

                    SearchResultCollection results = ds.FindAll();

                    foreach (SearchResult result in results)
                    {
                        using (DirectoryEntry uEntry = result.GetDirectoryEntry())
                        {
                            ADUserProperties aduser = new ADUserProperties();

                            aduser.cn = uEntry.Properties["cn"].Value.ToString();
                            aduser.displayName = uEntry.Properties["displayName"].Value == null ? "" : aduser.displayName = uEntry.Properties["displayName"].Value.ToString();
                            aduser.name = uEntry.Properties["name"].Value == null ? "" : aduser.name = uEntry.Properties["name"].Value.ToString();
                            aduser.profilePath = uEntry.Properties["profilePath"].Value == null ? "" : aduser.profilePath = uEntry.Properties["profilePath"].Value.ToString();
                            aduser.SamAccountName = uEntry.Properties["sAMAccountName"].Value == null ? "" : aduser.SamAccountName = uEntry.Properties["sAMAccountName"].Value.ToString();
                            aduser.userPrincipalName = uEntry.Properties["userPrincipalName"].Value == null ? "" : aduser.userPrincipalName = uEntry.Properties["userPrincipalName"].Value.ToString();
                            aduser.homeDirectory = uEntry.Properties["homeDirectory"].Value == null ? "" : aduser.homeDirectory = uEntry.Properties["homeDirectory"].Value.ToString();
                            aduser.mailNickname = uEntry.Properties["mailNickname"].Value == null ? "" : aduser.mailNickname = uEntry.Properties["mailNickname"].Value.ToString();
                            aduser.sn = uEntry.Properties["sn"].Value == null ? "" : aduser.sn = uEntry.Properties["sn"].Value.ToString();
                            Users.Add(aduser);
                        }
                    }
                }
            }

            Console.WriteLine("Users found: " + Users.Count);

            Console.WriteLine("Checking for duplicates surnames that may cause conflicts...");

            List<string> surnames = Users.Select(f => f.sn).ToList();
            List<string> duplicateSurnames = surnames.GroupBy(x => x).Where(g => g.Count() > 1).Select(y => y.Key).ToList();

            bool conflictsFound = false;
            foreach (string duplicateSurname in duplicateSurnames)
            {
                List<ADUserProperties> dupes = Users.Where(f => f.sn == duplicateSurname).ToList();

                if (dupes[0].name[0] == dupes[1].name[0])
                {
                    Users.Remove(dupes[0]);
                    Users.Remove(dupes[1]);
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("Username conflict detected: " + dupes[0].name + " AND " + dupes[1].name);
                    Console.WriteLine("These users will be skipped and should be done manually. They are written to duplicates.txt");
                    Console.ForegroundColor = ConsoleColor.White;
                    conflictsFound = true;
                    File.AppendAllText("duplicates.txt", "Skipped users: " + dupes[0].displayName + " and " + dupes[1].displayName + Environment.NewLine);
                }
            }

            if(!conflictsFound)
            {
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("No duplicate user conflicts have been found!");
                Console.ForegroundColor = ConsoleColor.White;
            }


            Console.WriteLine("We're all set! Press any key to get started...");
            Console.ReadKey(true);

            //Sort list alphabetically
            Users = Users.AsQueryable().OrderBy(f => f.displayName).ToList();

            int count = 0;

            foreach(ADUserProperties user in Users)
            {
                //Get login name, eg. j.smith
                string loginName = getLoginName(user.displayName);

                Console.WriteLine();
                Console.WriteLine("(" + count++ + "/" + Users.Count + ") The following properties will be changed to user: " + user.displayName);
                Console.WriteLine();

                //Create POST AD object (new settings)
                ADUserProperties aduser = new ADUserProperties();
                aduser.cn = loginName;
                aduser.displayName = user.displayName;
                aduser.mailNickname = loginName;
                aduser.name = loginName;
                aduser.SamAccountName = loginName;
                aduser.userPrincipalName = loginName + "@msa.nl";


                string dirProfileSource = "";
                string dirProfileDestination;
                bool hasProfileMap;
                bool hasUserMap;

                try
                {
                    //Find and determine the profile folder name
                    dirProfileSource = profileFolder + new DirectoryInfo(Directory.GetDirectories(profileFolder).Where(f => new DirectoryInfo(f).Name.ToLower() == user.SamAccountName.ToLower() + ".v2").FirstOrDefault()).Name.ToLower();
                    dirProfileDestination = profileFolder + loginName + ".V2";
                    hasProfileMap = true;
                }
                catch
                {
                    hasProfileMap = false;
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("WARNING: Profile folder not found or has already been renamed. Proceed?");
                    dirProfileDestination = profileFolder + loginName + ".V2";
                    Console.ForegroundColor = ConsoleColor.White;
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine("Select: Y / N");
                    Console.ForegroundColor = ConsoleColor.White;
                    Start:
                    var input = Console.ReadKey(true);
                    if (input.Key == ConsoleKey.N)
                    {
                        Console.ForegroundColor = ConsoleColor.Yellow;
                        Console.WriteLine();
                        Console.WriteLine("No changes have been made, moving on to next user.");
                        Console.ForegroundColor = ConsoleColor.White;
                        continue;
                    }
                    else if (input.Key == ConsoleKey.Y)
                    {
                        //Do nothing and proceed
                    }
                    else
                    {
                        goto Start;
                    }
                }

                string dirUserSource = "";
                string dirUserDestination = "";

                try
                {
                    //Find and determine the user folder name
                    dirUserSource = userFolder + new DirectoryInfo(Directory.GetDirectories(userFolder).Where(f => new DirectoryInfo(f).Name.ToLower() == user.SamAccountName.ToLower()).FirstOrDefault()).Name.ToLower();
                    dirUserDestination = userFolder + loginName;
                    hasUserMap = true;
                }
                catch
                {
                    hasUserMap = false;
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("WARNING: Userfolder not found or has already been renamed. Proceed?");
                    dirUserDestination = userFolder + loginName;
                    Console.ForegroundColor = ConsoleColor.White;
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine("Select: Y / N");
                    Console.ForegroundColor = ConsoleColor.White;
                    Start:
                    var input = Console.ReadKey(true);
                    if (input.Key == ConsoleKey.N)
                    {
                        Console.ForegroundColor = ConsoleColor.Yellow;
                        Console.WriteLine();
                        Console.WriteLine("No changes have been made, moving on to next user.");
                        Console.ForegroundColor = ConsoleColor.White;
                        continue;
                    }
                    else if (input.Key == ConsoleKey.Y)
                    {
                        //Do nothing and proceed
                    }
                    else
                    {
                        goto Start;
                    }
                }

                aduser.homeDirectory = dirUserDestination;
                aduser.profilePath = dirProfileDestination.Remove(dirProfileDestination.Length - 3);
                
                Console.WriteLine("{0,-20}{1,-50}{2,-10}{3,-1}", "Property", "From", "", "To");
                if(user.SamAccountName != aduser.SamAccountName)
                    Console.WriteLine("{0,-20}{1,-50}{2,-10}{3,-1}", "sAMAccountName:" , user.SamAccountName , " ========> " , aduser.SamAccountName);

                if (user.cn != aduser.cn)
                    Console.WriteLine("{0,-20}{1,-50}{2,-10}{3,-1}", "cn:" , user.cn , " ========> " , aduser.cn, "{0,-40}{1,-40}{2,-40}{3,-40}");

                if (user.mailNickname != aduser.mailNickname)
                    Console.WriteLine("{0,-20}{1,-50}{2,-10}{3,-1}", "mailNickname:" , user.mailNickname , " ========> " , aduser.mailNickname);

                if (user.userPrincipalName != aduser.userPrincipalName)
                    Console.WriteLine("{0,-20}{1,-50}{2,-10}{3,-1}", "userPrincipalName:" , user.userPrincipalName , " ========> " , aduser.userPrincipalName);

                if (user.name != aduser.name)
                    Console.WriteLine("{0,-20}{1,-50}{2,-10}{3,-1}", "name:" , user.name , " ========> " , aduser.name);

                if (hasProfileMap)
                    if(user.profilePath != aduser.profilePath)
                        Console.WriteLine("{0,-20}{1,-50}{2,-10}{3,-1}", "profilePath:" , user.profilePath , " ========> " , aduser.profilePath);

                if (hasUserMap)
                    if(user.homeDirectory != aduser.homeDirectory)
                        Console.WriteLine("{0,-20}{1,-50}{2,-10}{3,-1}", "homeDirectory:" , user.homeDirectory , " ========> " , aduser.homeDirectory);

                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine("Commit changes?: Y / N");
                Console.ForegroundColor = ConsoleColor.White;
                Start2:
                var input2 = Console.ReadKey(true);
                if (input2.Key == ConsoleKey.N)
                {
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.WriteLine();
                    Console.WriteLine("No changes have been made, moving on to next user.");
                    Console.ForegroundColor = ConsoleColor.White;
                }
                else if (input2.Key == ConsoleKey.Y)
                {
                    //rename directories
                    if (hasUserMap)
                    {
                        if(dirUserSource != dirUserDestination)
                            Directory.Move(dirUserSource, dirUserDestination);
                    }

                    if (hasProfileMap)
                    {
                        if(dirProfileSource != dirProfileDestination)
                            Directory.Move(dirProfileSource, dirProfileDestination);
                    }

                    //Commit changes in AD
                    setUserProperties(user.SamAccountName, aduser);

                    Console.ForegroundColor = ConsoleColor.DarkGreen;
                    Console.WriteLine();
                    Console.WriteLine("Changed saved, moving on to the next user.");
                    Console.ForegroundColor = ConsoleColor.White;
                }
                else
                {
                    goto Start2;
                }
                Console.ReadKey(true);
            }

            Console.ForegroundColor = ConsoleColor.Magenta;
            Console.WriteLine("All selected users have been handled, press any key to exit the application...");
            ForFun.End();
            Console.ReadKey(true);
        }


        public static string getLoginName(string displayName)
        {
            using (DirectoryEntry de = new DirectoryEntry(LDAP_scope))
            {
                using (DirectorySearcher ds = new DirectorySearcher(de))
                {
                    ds.Filter = string.Format("(&(objectcategory=user)({0}={1}))", "displayName", displayName);

                    SearchResult result = ds.FindOne();

                    using (DirectoryEntry uEntry = result.GetDirectoryEntry())
                    {
                        return (uEntry.Properties["givenName"].Value.ToString()[0] + "." + uEntry.Properties["sn"].Value.ToString()).ToLower();
                    }
                }
            }
        }

        public static void setUserProperties(string sAMAccountName, ADUserProperties Properties)
        {
            using (DirectoryEntry de = new DirectoryEntry(LDAP_scope))
            {
                using (DirectorySearcher ds = new DirectorySearcher(de))
                {
                    ds.Filter = string.Format("(&(objectcategory=user)({0}={1}))", "sAMAccountName", sAMAccountName);

                    SearchResult result = ds.FindOne();

                    using (DirectoryEntry uEntry = result.GetDirectoryEntry())
                    {
                        //Change values in AD
                        uEntry.Properties["sAMAccountName"].Value = Properties.SamAccountName;
                        uEntry.Properties["mailNickname"].Value = Properties.mailNickname;
                        uEntry.Properties["userPrincipalName"].Value = Properties.userPrincipalName;

                        //Path values
                        uEntry.Properties["homeDirectory"].Value = Properties.homeDirectory;
                        uEntry.Properties["profilePath"].Value = Properties.profilePath;
                        uEntry.CommitChanges();

                        uEntry.Rename("CN=" + Properties.name);
                        uEntry.Close();
                    }
                }
            }
        }
    }

    public class ADUserProperties
    {
        public string SamAccountName { get; set; }
        public string cn { get; set; }
        public string displayName { get; set; }
        public string mailNickname { get; set; }
        public string userPrincipalName { get; set; }
        public string name { get; set; }

        public string profilePath { get; set; }
        public string homeDirectory { get; set; }

        //Only to check for dupes
        public string sn { get; set; }
    }

    public class ForFun
    {
        private static string theEnd = @".___________. __    __   _______     _______ .__   __.  _______   __  
|           ||  |  |  | |   ____|   |   ____||  \ |  | |       \ |  | 
`---|  |----`|  |__|  | |  |__      |  |__   |   \|  | |  .--.  ||  | 
    |  |     |   __   | |   __|     |   __|  |  . `  | |  |  |  ||  | 
    |  |     |  |  |  | |  |____    |  |____ |  |\   | |  '--'  ||__| 
    |__|     |__|  |__| |_______|   |_______||__| \__| |_______/ (__) 
                                                                      ";
                                                                 

        public static async void End()
        {

            foreach(char c in theEnd)
            {
                Console.Write(c);
                await Task.Delay(1);
            }
        }
                   
    }
}
