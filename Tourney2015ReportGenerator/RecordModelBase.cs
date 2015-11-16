using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Tourney2015ReportGenerator
{
    public abstract class RecordModelBase
    {
        protected readonly RegistrationRecord _record;
        protected readonly TextInfo _textInfo = CultureInfo.CurrentCulture.TextInfo;
        public RecordModelBase(RegistrationRecord record)
        {
            _record = record;
        }

        public string FirstName { get { return _textInfo.ToTitleCase(_record.FirstName); } }
        public string LastName { get { return _textInfo.ToTitleCase(_record.LastName); } }
        public string FullName { get { return FirstName + " " + LastName; } }

        public abstract string[] TableRowData();
    }
}
