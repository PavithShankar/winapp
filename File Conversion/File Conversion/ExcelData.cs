using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace File_Conversion
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
        [DisplayName("From To Text Actor")]
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
                if (string.IsNullOrEmpty(actor))
                {
                    actor = string.Empty;
                    return actor;
                }
                else
                {

                    return actor;
                }
            
            }
            set
            {
                    actor = value;
             
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

