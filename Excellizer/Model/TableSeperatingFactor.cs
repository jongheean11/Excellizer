using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excellizer.Model
{
    class TableSeperatingFactor
    {
        private int idxRow;
        private int idxCol;

        int init_x
        {
            get;
            set;
        }
        int init_y
        {
            get;
            set;
        }
        int rowspan
        {
            get;
            set;
        }
        int colspan
        {
            get;
            set;
        }
        
        public TableSeperatingFactor(int init_x, int init_y, int rowspan, int colspan)
        {
            // TODO: Complete member initialization
            this.init_x = init_x;
            this.init_y = init_y;
            this.rowspan = rowspan;
            this.colspan = colspan;
        }
    }
}
