using System;
using System.Globalization;
using System.Linq;

namespace Tourney2015ReportGenerator
{
    public class Competitor : RecordModelBase
    {
        public Competitor(RegistrationRecord record) : base(record)
        { }

        public string SchoolName { get { return _textInfo.ToTitleCase(_record.SchoolName); } }
        public int Age { get { return int.Parse(_record.Age); } }

        public decimal Weight
        {
            get
            {
                var weight = _record.Weight.ToLower().Replace("s", "");
                if(weight.Contains("kg"))
                {
                    weight = weight.Replace("kg.", "").Replace("kg", "");
                    decimal kg = 0.0m;
                    if (decimal.TryParse(weight, out kg))
                        return ConverKgToLb(kg);
                    throw new ArgumentException($"Invalid weight: {_record.Weight}");
                }

                weight = weight.Replace("lb.", "").Replace("lb", "");
                decimal lb = 0.0m;
                if (decimal.TryParse(weight, out lb))
                    return lb;
                throw new ArgumentException($"Invalid weight: {_record.Weight}");
            }
        }

        protected static decimal ConverKgToLb(decimal kg)
        {
            return (kg * 2.2046226m);
        }

        public string Gender { get { return _record.Gender; } }

        public string Rank { get { return CalculateRank(); } }

        public AgeDivision AgeDivision { get { return CalculateAgeDivision(); } }

        public bool IsSparring { get { return CalculateIsSparring(); } }
        public bool IncludeInExtraRankDivision { get { return CalculateIncludeInExtraRankDivision(); } }
        public bool IncludeInExtraAgeDivision { get { return CalculateIncludeInExtraAgeDivision(); } }
        public bool IsSpecial { get { return IncludeInExtraRankDivision || IncludeInExtraAgeDivision; } }
        public bool IsForms { get { return CalculateIsForms(); } }

        public string NameBadgeEvent { get { return CalculateNameBadgeEvent(); } }

        public bool IsIncludedInRank(string rank)
        {
            if (rank == Rank)
                return true;

            if (IncludeInExtraRankDivision)
            {
                var normalIndex = Array.IndexOf(EventInfo.Ranks, Rank);
                if (EventInfo.Ranks.Length > normalIndex + 1)
                {
                    return EventInfo.Ranks[normalIndex + 1] == rank;
                }
            }
            return false;
        }

        public bool IsIncludedInAgeDivision(AgeDivision ageDivision)
        {
            if(ageDivision == AgeDivision)
                return true;

            if (IncludeInExtraAgeDivision)
            {
                var normalIndex = Array.IndexOf(EventInfo.AgeDivisions, AgeDivision);
                if (EventInfo.AgeDivisions.Length > normalIndex + 1)
                {
                    return EventInfo.AgeDivisions[normalIndex + 1] == ageDivision;
                }
            }
            return false;
        }

        protected string CalculateRank()
        {
            var ticketType = _record.TicketType;
            if (ticketType == EventInfo.SingleElimTicket)
                return _record.Rank;
            else if (ticketType == EventInfo.SingleAndDoubleElimTicket)
                return EventInfo.AdvancedRank;
            else
                return EventInfo.BlackBeltRank;
        }

        protected AgeDivision CalculateAgeDivision()
        {
            return EventInfo.AgeDivisions.FirstOrDefault(x =>
                x.StartAge <= Age && x.EndAge >= Age);
        }

        protected bool CalculateIsSparring()
        {
            return true;
        }

        protected bool CalculateIncludeInExtraAgeDivision()
        {
            var ticketType = _record.TicketType;
            switch (ticketType)
            {
                case EventInfo.DoubleElim2EventTicket:
                    return true;
                default:
                    return false;
            }
        }

        protected bool CalculateIncludeInExtraRankDivision()
        {
            var ticketType = _record.TicketType;
            switch (ticketType)
            {
                case EventInfo.SingleAndDoubleElimTicket:
                    return true;
                default:
                    return false;
            }
        }

        protected bool CalculateIsForms()
        {
            return false;
        }

        protected string CalculateNameBadgeEvent()
        {
            if (_record.TicketType == EventInfo.SingleElimTicket)
                return "Single-Elim Sparring";
            if(_record.TicketType == EventInfo.DoubleElim1EventTicket)
                return "Double-Elim Sparring";
            if (_record.TicketType == EventInfo.SingleAndDoubleElimTicket)
                return "Single and Double-Elim Sparring";
            if (_record.TicketType == EventInfo.DoubleElim2EventTicket)
                return "Double-Elim Sparring (2x)";
            return "";
        }

        public override string[] TableRowData()
        {
            return new string[]
            {
                FirstName,
                LastName + (IsSpecial ? "***" : ""),
                Weight.ToString(),
                SchoolName
            };
        }

        public string[] TableFullRowData()
        {
            return new string[]
            {
                FirstName,
                LastName + (IsSpecial ? "***" : ""),
                Gender,
                Age.ToString(),
                Weight.ToString(),
                Rank,
                SchoolName,
                NameBadgeEvent
            };
        }

        public static string[] TableColumnNames()
        {
            return new string[]
            {
                "First Name",
                "Last Name",
                "Weight",
                "School"
            };
        }

        public static string[] TableFullColumnNames()
        {
            return new string[]
            {
                "First Name",
                "Last Name",
                "Gender",
                "Age",
                "Weight",
                "Rank",
                "School",
                "Events",
            };   
        }  
    }      
}          
           
           