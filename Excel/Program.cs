using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Aurora.IO.Excel.Extensions;
using Excel.Models.Horizontal;
using Excel.Models.Vertical;
using Microsoft.Win32;
using OfficeOpenXml;

namespace Excel
{
    internal class Program
    {
        public static void Main(string[] args)
        {
            // Grab demo file
            using (var excel = new ExcelPackage(
                new FileInfo(
                    Path.Combine(new DirectoryInfo(Directory.GetCurrentDirectory()).Parent.Parent.Parent.FullName,
                        "DemoFile",
                        "ExcelDemo.xlsx")))) // If missing replace FileInfo constructor parameter with your own file
            {
                // Horizontal
                Console.WriteLine("Horizontal mapping");

                // Get sheet with horizontal mapping
                var sheet = excel.Workbook.Worksheets.First();
                
                // Get list of teams based on automatically mapping of the header row
                Console.WriteLine("List of teams based on automatically mapping of the header row");
                var teams = sheet.GetRecords<Team>();
                foreach (var team in teams)
                {
                    Console.WriteLine($"{team.Name} - {team.FoundationYear} - {team.Titles}");
                }
                
                // Get specific record from sheet based on automatically mapping the header row
                Console.WriteLine("Team based on automatically mapping of the header row");
                var teamRec = sheet.GetRecord<Team>(2);
                Console.WriteLine($"{teamRec.Name} - {teamRec.FoundationYear} - {teamRec.Titles}");
                
                // Remove HeaderRow
                sheet.DeleteRow(1);
                
                // Get list of teams based on mapping using attributes
                Console.WriteLine("List of teams based on mapping using attributes");
                var teamsAttr = sheet.GetRecords<TeamAttributes>();
                foreach (var team in teamsAttr)
                {
                    Console.WriteLine($"{team.Name} - {team.FoundationYear} - {team.Titles}");
                }
                
                // Get specific record from sheet based on mapping using attributes
                Console.WriteLine("Team based on mapping using attributes");
                var teamAttr = sheet.GetRecord<TeamAttributes>(1);
                Console.WriteLine($"{teamAttr.Name} - {teamAttr.FoundationYear} - {teamAttr.Titles}");
                
                // Get list of teams based on user created map
                Console.WriteLine("List of teams based on user created map");
                var teamsMap = sheet.GetRecords(TeamMap.Create());
                foreach (var team in teamsMap)
                {
                    Console.WriteLine($"{team.Name} - {team.FoundationYear} - {team.Titles}");
                }
                
                // Get specific record from sheet based on user created map
                Console.WriteLine("Team based on user created map");
                var teamMap = sheet.GetRecord<Team>(1, TeamMap.Create());
                Console.WriteLine($"{teamMap.Name} - {teamMap.FoundationYear} - {teamMap.Titles}");

                // Vertical
                Console.WriteLine("Vertical mapping");

                // Get sheet with vertical mapping
                var vsheet = excel.Workbook.Worksheets.Skip(1).Take(1).First();

                // Get list of teams based on automatically mapping of the header row
                Console.WriteLine("List of teams based on automatically mapping of the header row");
                var vteams = vsheet.GetRecords<VTeam>();
                foreach (var vteam in vteams)
                {
                    Console.WriteLine($"{vteam.Name} - {vteam.FoundationYear} - {vteam.Titles}");
                }

                // Get specific record from sheet based on automatically mapping the header row
                Console.WriteLine("Team based on automatically mapping of the header row");
                var vteamRec = vsheet.GetRecord<VTeam>(2);
                Console.WriteLine($"{vteamRec.Name} - {vteamRec.FoundationYear} - {vteamRec.Titles}");

                // Remove HeaderRow
                vsheet.DeleteColumn(1);

                // Get list of teams based on mapping using attributes
                Console.WriteLine("List of teams based on mapping using attributes");
                var vteamsAttr = vsheet.GetRecords<VTeamAttributes>();
                foreach (var vteam in vteamsAttr)
                {
                    Console.WriteLine($"{vteam.Name} - {vteam.FoundationYear} - {vteam.Titles}");
                }

                // Get specific record from sheet based on mapping using attributes
                Console.WriteLine("Team based on mapping using attributes");
                var vteamAttr = vsheet.GetRecord<VTeamAttributes>(1);
                Console.WriteLine($"{vteamAttr.Name} - {vteamAttr.FoundationYear} - {vteamAttr.Titles}");

                // Get list of teams based on user created map
                Console.WriteLine("List of teams based on user created map");
                var vteamsMap = vsheet.GetRecords(VTeamMap.Create());
                foreach (var vteam in vteamsMap)
                {
                    Console.WriteLine($"{vteam.Name} - {vteam.FoundationYear} - {vteam.Titles}");
                }

                // Get specific record from sheet based on user created map
                Console.WriteLine("Team based on user created map");
                var vteamMap = vsheet.GetRecord<VTeam>(1, VTeamMap.Create());
                Console.WriteLine($"{vteamMap.Name} - {vteamMap.FoundationYear} - {vteamMap.Titles}");

                Console.Read();
            }
        }
    }
}