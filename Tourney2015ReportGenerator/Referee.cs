using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Tourney2015ReportGenerator
{
    public class Referee : RecordModelBase
    {
        public Referee(RegistrationRecord record) : base(record)
        { }

        public string RefereeLevel { get { return _record.RefereeLevel; } }

        public override string[] TableRowData()
        {
            return new string[]
            {
                FullName,
                RefereeLevel
            };
        }

        public static string[] TableColumnNames()
        {
            return new string[]
            {
                "Name",
                "Referee Level"
            };
        }
    }
}
