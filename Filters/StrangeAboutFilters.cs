using System;
using System.Collections.Generic;
using System.Text;
using System.Collections;

using EducServLib;

namespace Priem
{
    public class StrangeAboutFilters
    {
        //строит строку фильтров для списка с фильтрами по коллекции из формы фильтров
        public static string BuildFilterWithCollection(List<iFilter> list)
        {
            IEnumerator ienum = list.GetEnumerator();
            string res = string.Format(" AND ({0}) ", BuildFilter(ienum));
            return res;
        }

        public static string BuildFilter(IEnumerator ienum)
        {
            string res = " 1=1 ";

            while (ienum.MoveNext())
            {
                object obj = ienum.Current;
                if (obj is Filter)
                    res += string.Format(" AND {0} ", (obj as Filter).GetFilter());
                else if (obj is Or)
                    res = string.Format("({0}) OR ({1})", res, BuildFilter(ienum));
                else if (obj is LeftBracket)
                    res = string.Format("{0} AND ({1})", res, BuildFilter(ienum));
                else if (obj is RightBracket)
                    break;
            }

            return res;
        }
    }
}
