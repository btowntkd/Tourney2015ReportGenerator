using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Tourney2015ReportGenerator
{
    public static class EventInfo
    {
        public const string RefereeTicket = "Referee";
        public const string CoachTicket = "Coach";
        public const string SingleElimTicket = "Single-Elimination Sparring";
        public const string SingleAndDoubleElimTicket = "Single-Elimination and Double-Elimination (Red-Belt Only)";
        public const string DoubleElim1EventTicket = "Double-Elimination Sparring";
        public const string DoubleElim2EventTicket = "Double-Elimination Sparring (2 Divisions)";
        public const string SpectatorAdultTicket = "Spectator (Adult, age 12+)";
        public const string SpectatorChildTicket = "Spectator (Child, age 6-12)";

        private static readonly AgeDivision[] _ageDivisions = 
        {
            new AgeDivision("< 6", 0, 5),
            new AgeDivision("6-7", 6, 7),
            new AgeDivision("8-9", 8, 9),
            new AgeDivision("10-11", 10, 11),
            new AgeDivision("12-14", 12, 14),
            new AgeDivision("15-17", 15, 17),
            new AgeDivision("18-32", 18, 32),
            new AgeDivision("> 32", 32, 9999),
        };

        private static readonly string[] _ranks =
        {
            BeginnerRank,
            Intermediate1Rank,
            Intermediate2Rank,
            AdvancedRank,
            BlackBeltRank
        };

        public const string BeginnerRank = "Beginner (10th geup - 8th geup)";
        public const string Intermediate1Rank = "Intermediate I (7th geup - 5th geup)";
        public const string Intermediate2Rank = "Intermediate II (4th geup - 3rd geup)";
        public const string AdvancedRank = "Advanced (2nd geup - 1st geup)";
        public const string BlackBeltRank = "Black Belt";

        public static AgeDivision[] AgeDivisions { get { return _ageDivisions; } }
        public static string[] Ranks { get { return _ranks; } }
    }

    public class AgeDivision
    {
        public int StartAge { get; private set; }
        public int EndAge { get; private set; }
        public string Name { get; private set; }

        public AgeDivision(string name, int start, int end)
        {
            Name = name;
            StartAge = start;
            EndAge = end;
        }
    }
}
