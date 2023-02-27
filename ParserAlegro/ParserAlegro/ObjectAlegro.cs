using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ParserAlegro
{
    public class ObjectAlegro
    {
        public string Title { get; set; }
        public string NumberLot { get; set; }
        public List<string> Photos { get; set; } = new();
        public string Price { get; set; }
        public string CatalogNumber { get; set; }
        public string CatalogNumber_2 { get; set; }
        public string CatalogNumber_3 { get; set; }
        
        public ObjectAlegro()
        {

        }

        public override string ToString()
        {
            return $"Title -> {Title}. NumberLot -> {NumberLot}. Price -> {Price}. CatalogNumber -> {CatalogNumber}";
        }
    }
}
