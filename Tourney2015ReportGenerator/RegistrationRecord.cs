using FileHelpers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Tourney2015ReportGenerator
{
    [DelimitedRecord(",")]
    [IgnoreFirst(1)]
    public class RegistrationRecord
    {
        [FieldQuoted('"', QuoteMode.OptionalForBoth)]
        public string OrderNumber;
        [FieldQuoted('"', QuoteMode.OptionalForBoth)]
        public string FirstName;
        [FieldQuoted('"', QuoteMode.OptionalForBoth)]
        public string LastName;
        [FieldQuoted('"', QuoteMode.OptionalForBoth)]
        public string Email;
        [FieldQuoted('"', QuoteMode.OptionalForBoth)]
        public string TicketType;
        [FieldQuoted('"', QuoteMode.OptionalForBoth)]
        public string Gender;
        [FieldQuoted('"', QuoteMode.OptionalForBoth)]
        public string Age;
        [FieldQuoted('"', QuoteMode.OptionalForBoth)]
        public string Rank;
        [FieldQuoted('"', QuoteMode.OptionalForBoth)]
        public string Weight;
        [FieldQuoted('"', QuoteMode.OptionalForBoth)]
        public string SchoolName;
        [FieldQuoted('"', QuoteMode.OptionalForBoth)]
        public string InstructorName;
        [FieldQuoted('"', QuoteMode.OptionalForBoth)]
        public string RefereeLevel;
        [FieldQuoted('"', QuoteMode.OptionalForBoth)]
        public string SelectOneDoubleElim;
        [FieldQuoted('"', QuoteMode.OptionalForBoth)]
        public string SelectTwoDoubleElim;
        [FieldQuoted('"', QuoteMode.OptionalForBoth)]
        public string SelectSingleElimEvents;
        [FieldQuoted('"', QuoteMode.OptionalForBoth)]
        public string LiabilityWaiver;
    }
}
