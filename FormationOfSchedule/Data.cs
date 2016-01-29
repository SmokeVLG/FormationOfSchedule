using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace FormationOfSchedule
{
    class Data
    {

        public Data()
		{ }

		public Data(string contrType, string contrName, string KSSS, string contrcode)
		{
            _contrType = contrType;
            _contrName = contrName;
            _KSSS = KSSS;
            _contrcode = contrcode;
		}
		private string _contrType;
        public string contrType
		{
			get
			{
                return _contrType;
			}
			set
			{
                _contrType = value;
			}
		}

        private string _contrName;
        public string contrName
		{
			get
			{
                return _contrName;
			}
			set
			{
                _contrName = value;
			}
		}

        private string _KSSS;
        public string KSSS
		{
			get
			{
                return _KSSS;
			}
			set
			{
                _KSSS = value;
			}
		}

        private string _contrcode;
        public string contrcode
		{
			get
			{
                return _contrcode;
			}
			set
			{
                _contrcode = value;
			}
		}

    }
}
