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

        public decimal Weight { get { return decimal.Parse(_record.Weight.ToLower().Replace("lbs", "")); } }

        public string Gender { get { return _record.Gender; } }

        public string Rank { get { return CalculateRank(); } }

        public AgeDivision AgeDivision { get { return CalculateAgeDivision(); } }

        public bool IsSparring { get { return CalculateIsSparring(); } }
        public bool IsForms { get { return CalculateIsForms(); } }

        public string NameBadgeEvent { get { return CalculateNameBadgeEvent(); } }

        protected string CalculateRank()
        {
            var ticketType = _record.TicketType;
            if (ticketType == EventInfo.SingleElimTicket)
                return _record.Rank;
            else
                return "Black Belt";
        }

        protected AgeDivision CalculateAgeDivision()
        {
            return EventInfo.AgeDivisions.FirstOrDefault(x =>
                x.StartAge <= Age && x.EndAge >= Age);
        }

        protected bool CalculateIsSparring()
        {
            var ticketType = _record.TicketType;
            switch (ticketType)
            {
                case EventInfo.SingleElimTicket:
                    {
                        var selectedEvent = _record.SelectSingleElimEvents;
                        return (selectedEvent == EventInfo.SingleElimSparringOnly
                            || selectedEvent == EventInfo.SingleElimFormsAndSparring);
                    }
                case EventInfo.DoubleElim1EventTicket:
                    {
                        var selectedEvent = _record.SelectOneDoubleElim;
                        return selectedEvent == EventInfo.DoubleElimSparringOnly;
                    }
                case EventInfo.DoubleElim2EventTicket:
                    {
                        var selectedEvent = _record.SelectTwoDoubleElim;
                        return (selectedEvent == EventInfo.DoubleElimFormsAndSparring
                            || selectedEvent == EventInfo.DoubleElimDoubleSparring);
                    }
                case EventInfo.DoubleElim3EventTicket:
                    return true;
                default:
                    return false;
            }
        }

        protected bool CalculateIsForms()
        {
            var ticketType = _record.TicketType;
            switch (ticketType)
            {
                case EventInfo.SingleElimTicket:
                    {
                        var selectedEvent = _record.SelectSingleElimEvents;
                        return (selectedEvent == EventInfo.SingleElimFormsOnly
                            || selectedEvent == EventInfo.SingleElimFormsAndSparring);
                    }
                case EventInfo.DoubleElim1EventTicket:
                    {
                        var selectedEvent = _record.SelectOneDoubleElim;
                        return selectedEvent == EventInfo.DoubleElimFormsOnly;
                    }
                case EventInfo.DoubleElim2EventTicket:
                    {
                        var selectedEvent = _record.SelectTwoDoubleElim;
                        return selectedEvent == EventInfo.DoubleElimFormsAndSparring;
                    }
                case EventInfo.DoubleElim3EventTicket:
                    return true;
                default:
                    return false;
            }
        }

        protected string CalculateNameBadgeEvent()
        {
            if (_record.SelectSingleElimEvents == EventInfo.SingleElimSparringOnly
                || _record.SelectOneDoubleElim == EventInfo.DoubleElimSparringOnly)
                return "Sparring";
            if (_record.SelectSingleElimEvents == EventInfo.SingleElimFormsOnly
                || _record.SelectOneDoubleElim == EventInfo.DoubleElimFormsOnly)
                return "Forms";
            if (_record.SelectSingleElimEvents == EventInfo.SingleElimFormsAndSparring
                || _record.SelectTwoDoubleElim == EventInfo.DoubleElimFormsAndSparring)
                return "Forms & Sparring";
            if (_record.SelectTwoDoubleElim == EventInfo.DoubleElimDoubleSparring)
                return "2x Sparring";

            return "Forms & 2x Sparring";
        }

        public override string[] TableRowData()
        {
            return new string[]
            {
                FullName,
                Weight.ToString(),
                SchoolName
            };
        }

        public string[] TableFullRowData()
        {
            return new string[]
            {
                FullName,
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
                "Name",
                "Weight",
                "School"
            };
        }

        public static string[] TableFullColumnNames()
        {
            return new string[]
            {
                "Name",
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
           
           