using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Tourney2015ReportGenerator
{
    public class Spectator : RecordModelBase
    {
        private readonly IEnumerable<RegistrationRecord> _records;
        public Spectator(IEnumerable<RegistrationRecord> records) : base(records.FirstOrDefault())
        {
            _records = records;
        }

        public int NumAdults { get { return _records.Count(x => x.TicketType == EventInfo.SpectatorAdultTicket); } }
        public int NumChildren { get { return _records.Count(x => x.TicketType == EventInfo.SpectatorChildTicket); } }

        public override string[] TableRowData()
        {
            return new string[]
            {
                LastName + ", " + FirstName,
                NumAdults.ToString(),
                NumChildren.ToString()
            };
        }

        public static string[] TableColumnNames()
        {
            return new string[]
            {
                "Name",
                "Adult Tickets",
                "Child Tickets"
            };
        }
    }
}
