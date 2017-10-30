using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Tourney2015ReportGenerator
{
    public class RegistrationReporter
    {
        private readonly RegistrationRecord[] _records;
        public RegistrationReporter(IEnumerable<RegistrationRecord> records)
        {
            _records = records.ToArray();
        }

        public Competitor[] GetCompetitors()
        {
            return _records.Where(x => IsCompetitor(x))
                .Select(x => new Competitor(x))
                .ToArray();
        }

        public Spectator[] GetSpectators()
        {
            return _records.Where(x => IsSpectator(x))
                .GroupBy(x => new { x.FirstName, x.LastName })
                .Select(x => new Spectator(x))
                .ToArray();
        }

        public Referee[] GetReferees()
        {
            return _records.Where(x => IsReferee(x))
                .Select(x => new Referee(x))
                .ToArray();
        }

        public Coach[] GetCoaches()
        {
            return _records.Where(x => IsCoach(x))
                .Select(x => new Coach(x))
                .ToArray();
        }

        public bool IsSpectator(RegistrationRecord record)
        {
            return record.TicketType == EventInfo.SpectatorAdultTicket
                || record.TicketType == EventInfo.SpectatorChildTicket;
        }

        protected bool IsCompetitor(RegistrationRecord record)
        {
            return record.TicketType == EventInfo.SingleElimTicket
                || record.TicketType == EventInfo.SingleAndDoubleElimTicket
                || record.TicketType == EventInfo.DoubleElim1EventTicket
                || record.TicketType == EventInfo.DoubleElim2EventTicket;
        }

        protected bool IsReferee(RegistrationRecord record)
        {
            return record.TicketType == EventInfo.RefereeTicket;
        }

        protected bool IsCoach(RegistrationRecord record)
        {
            return record.TicketType == EventInfo.CoachTicket;
        }
    }
}
