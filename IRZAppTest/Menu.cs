using System;
using System.Collections.Generic;
using System.Text;

namespace IRZAppTest
{
    public class Menu //api.stackexchange.com response structure in Json format
    {
        public Item[] items { get; set; }
    }
    public class Item
    {
        public Owner owner { get; set; }
        public bool is_answered { get; set; }
        public string link { get; set; }
        public string title { get; set; }
    }
    public class Owner
    {
        public string display_name { get; set; }
        public string link { get; set; }
    }
}
