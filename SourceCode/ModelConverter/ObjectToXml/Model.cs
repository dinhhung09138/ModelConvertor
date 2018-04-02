using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Serialization;

namespace ModelConverter.ObjectToXml
{
    [XmlRoot(ElementName = "Employee", Namespace = "")]
    public class Model
    {

        public string ID { get; set; } = "";

        public string Name { get; set; } = "";

        [XmlIgnore]
        public string Address { get; set; } = "";

        public DateTime BirthDay { get; set; } = DateTime.Now;

        [XmlElement(ElementName = "YouAge")]
        public int Age { get; set; } = 0;

        [XmlElement(ElementName = "Phone")]
        public string Phone { get; set; } = "";

        [XmlElement(ElementName = "Details")]
        public List<ModelDetail> List { get; set; } = new List<ModelDetail>();

    }

    public class ModelDetail
    {
        public int DetailID { get; set; } = 0;
        public int ProductID { get; set; } = 0;
        public DateTime CreateDate { get; set; } = DateTime.Now;
    }
}
