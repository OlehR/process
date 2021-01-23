using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelApp = Microsoft.Office.Interop.Excel;

namespace Report
{
    public class сRequest
    {
        /// <summary>
        /// SQL чи MDX запит
        /// </summary>
        public string Request { get; set; }

        public eClient Client { get; set; }
        /// <summary>
        /// Сторінка куда записати результат
        /// </summary>
        public ExcelApp.Worksheet Sheet { get; set; }
        /// <summary>
        /// Рядок з відки почати вставляти дані
        /// </summary>
        public int Row { get; set; }
        /// <summary>
        /// Колонка з відки почати вставляти дані
        /// </summary>
        public int Column { get; set; }

        public bool IsHead = true;

    }
}
