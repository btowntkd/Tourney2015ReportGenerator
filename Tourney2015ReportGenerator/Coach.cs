using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Tourney2015ReportGenerator
{
    public class Coach : RecordModelBase
    {
        public Coach(RegistrationRecord record) : base(record)
        { }

        public string SchoolName { get { return _textInfo.ToTitleCase(_record.SchoolName); } }

        public override string[] TableRowData()
        {
            return new string[]
            {
                FullName,
                SchoolName
            };
        }

        public static string[] TableColumnNames()
        {
            return new string[]
            {
                "Name",
                "School"
            };
        }
    }
}
