using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ComponentModel;
using System.Threading.Tasks;

namespace Excel_To_Json_Converter_WinForms_
{
    class ExcelData
    {
        #region Members
        private int from;
        private int to;
        private string text;
        private string actor;
        #endregion

        #region Prperties
        [DisplayNameAttribute("From To Text Actor")]
        public int From
        {
            get
            {
                return from;
            }
            set
            {
               
               from = value;
            }
        }

        public int To
        {
            get
            {
                return to;
            }
            set
            {
                to = value;
            }
        }
        public string Text
        {
            get
            {
                return text;
            }
            set
            {
                text = value;
            }

        }
        public string Actor
        {
            get
            {

                return actor;
            }
            set
            {
                if (string.IsNullOrEmpty(Actor))
                {
                    Actor = "";

                }
                else
                {
                    actor = value;
                }

            }

        }
        #endregion

        #region Intialization
        public ExcelData()
        {
        }

        #endregion

    }
}
