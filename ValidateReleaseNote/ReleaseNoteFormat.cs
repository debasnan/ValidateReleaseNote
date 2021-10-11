using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ValidateReleaseNote
{
    public class ReleaseNoteFormat
    {
        public string LineNumber { get; set; }
        public string CROrDefectReference { get; set; }
        public string FileName { get; set; }
        public string ChangeType { get; set; }
        public string GitCheckInReference { get; set; }
        public string Developer { get; set; }
    }
}
