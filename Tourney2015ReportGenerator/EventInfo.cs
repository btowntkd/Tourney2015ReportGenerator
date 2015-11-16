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
        public const string SingleElimTicket = "Single-Elimination Events";
        public const string DoubleElim1EventTicket = "Double-Elimination and/or Sport Poomsae (1 Event)";
        public const string DoubleElim2EventTicket = "Double-Elimination and/or Sport Poomsae (2 Events)";
        public const string DoubleElim3EventTicket = "Double-Elimination and/or Sport Poomsae (3 Events)";
        public const string SpectatorAdultTicket = "Spectator (Adult, age 12+)";
        public const string SpectatorChildTicket = "Spectator (Child, age 6-12)";


        public const string SingleElimFormsAndSparring = "Forms and Sparring";
        public const string SingleElimFormsOnly = "Forms only";
        public const string SingleElimSparringOnly = "Sparring only";

        public const string DoubleElimFormsAndSparring = "Forms and Sparring";
        public const string DoubleElimFormsOnly = "Forms only";
        public const string DoubleElimSparringOnly = "Sparring only";
        public const string DoubleElimDoubleSparring = "2x Sparring Events (two age divisions)";

        private static readonly AgeDivision[] _ageDivisions = 
        {
            new AgeDivision("< 6", 0, 5),
            new AgeDivision("6-7", 6, 7),
            new AgeDivision("8-9", 8, 9),
            new AgeDivision("10-11", 10, 11),
            new AgeDivision("12-14", 12, 14),
            new AgeDivision("15-17", 15, 17),
            new AgeDivision("18-32", 18, 32),
            new AgeDivision("> 32", 32, 99),
        };

        public const string BlackBeltRank = "Black Belt";

        public static AgeDivision[] AgeDivisions { get { return _ageDivisions; } }
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
