using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReadExcel
{
    class ItemData
    {
        public class itemData
        {
            public string _item;
            public object _itemData;

            public itemData(string item, object data)
            {
                _item = item;
                _itemData = data;
            }

            public override string ToString()
            {
                return _item;
            }
        }
    }
}
