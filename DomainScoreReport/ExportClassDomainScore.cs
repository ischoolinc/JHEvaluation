using System.Collections.Generic;

namespace DomainScoreReport
{
    class ExportClassDomainScore
    {
        private readonly List<string> listClassIDs = new List<string>();

        public ExportClassDomainScore(List<string>listIDs)
        {
            listClassIDs = listIDs;
        }

        /// <summary>
        /// 成績單列印
        /// </summary>
        public void Export()
        {

        }
    }
}
