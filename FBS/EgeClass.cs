using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Priem
{
    //свидетельство ЕГЭ
    public class FBSEgeCert
    {
        private string _name;//номер свидетельства
        private string _tipograf;
        private string _year;
        private List<FBSEgeMark> _markList;//оценки

        //constructor
        public FBSEgeCert(string name, string tipograf, string year)
        {
            _name = name;
            _tipograf = tipograf;
            _year = year;
            _markList = new List<FBSEgeMark>();
        }

        //constructor
        public FBSEgeCert(string name, List<FBSEgeMark> markList)
        {
            _name = name;
            _markList = markList;
        }

        //property
        public string Name
        {
            get
            {
                return _name;
            }
        }

        public string Tipograf
        {
            get
            {
                return _tipograf;
            }
        }

        public string Year
        {
            get
            {
                return _year;
            }
        }
        //property
        public List<FBSEgeMark> Marks
        {
            get
            {
                return _markList;
            }
        }

        //add new mark in list
        public void AddMark(FBSEgeMark mark)
        {
            _markList.Add(mark);
        }
    }

    public class FBSEgeMark
    {
        private int _value;
        private string _examId;
        private bool _isApl;

        //constructor
        public FBSEgeMark(string sExamId, int iValue, bool sIsApl)
        {
            _examId = sExamId;
            _value = iValue;
            _isApl = sIsApl;
        }

        //property
        public string ExamId
        {
            get
            {
                return _examId;
            }
        }

        //property
        public int Value
        {
            get
            {
                return _value;
            }
        }

        //property
        public bool isApl
        {
            get
            {
                return _isApl;
            }
        }
    }    
}
