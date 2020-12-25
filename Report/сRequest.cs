using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
        public string Sheet { get; set; }
        /// <summary>
        /// Рядок з відки почати вставляти дані
        /// </summary>
        public int Row { get; set; }
        /// <summary>
        /// Колонка з відки почати вставляти дані
        /// </summary>
        public int Column { get; set; }

    }
}
