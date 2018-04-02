using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Serialization;

namespace ModelConverter.ObjectToXml
{
    public class Execute
    {

        public void Mail()
        {
            #region " [ Convert object to xml string ] "

            //Model _model = new Model()
            //{
            //    ID = "1",
            //    Name = "Hung",
            //    Address = "Go vap",
            //    Age = 23,
            //    BirthDay = DateTime.Now.AddYears(-27),
            //    Phone = "212",
            //    List = new List<ModelDetail>()
            //    {
            //        new ModelDetail() { DetailID = 1, ProductID = 1, CreateDate = DateTime.Now },
            //        new ModelDetail() { DetailID = 2, ProductID = 2, CreateDate = DateTime.Now.AddYears(1) }
            //    }
            //};
            //string _xml = ConvertToXML(_model);
            //_xml = FormatXmlContent(_xml);
            //FileHelper.WriteFile(Path.Combine(Environment.CurrentDirectory, "Employee.xml"), _xml, false);

            #endregion

            #region " [ Convert xml content to Object model ] "

            //string _xmlContent = FileHelper.ReadTextFile(Path.Combine(Environment.CurrentDirectory, "Employee.xml"));
            //Console.WriteLine("");
            //Console.WriteLine("");
            //Console.WriteLine(_xmlContent);
            //Model _model = ConvertToObject(_xmlContent, typeof(Model)) as Model;
            //Console.WriteLine("");
            //Console.WriteLine("");
            //Console.WriteLine("");

            #endregion
        }

        /// <summary>
        /// Convert to xml string from object model
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        public string ConvertToXML(object model)
        {
            StringWriter sw = new StringWriter();
            XmlTextWriter tw = null;
            try
            {
                XmlSerializerNamespaces _xSN = new XmlSerializerNamespaces();
                _xSN.Add("", "");

                XmlSerializer serializer = new XmlSerializer(model.GetType());
                tw = new XmlTextWriter(sw);
                serializer.Serialize(tw, model, _xSN);

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                Console.WriteLine("Finish write xml from object model");
                sw.Close();
                if (tw != null)
                {
                    tw.Close();
                }
            }
            return sw.ToString();
        }

        /// <summary>
        /// Convert to object model from xml string
        /// </summary>
        /// <param name="xml">xml string</param>
        /// <param name="objectType">object model type</param>
        /// <returns></returns>
        public object ConvertToObject(string xml, Type objectType)
        {
            StringReader _strReader = null;
            XmlSerializer _serializer = null;
            XmlTextReader _xmlReader = null;
            Object _obj = null;
            try
            {
                _strReader = new StringReader(xml);
                _serializer = new XmlSerializer(objectType);
                _xmlReader = new XmlTextReader(_strReader);
                _obj = _serializer.Deserialize(_xmlReader);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                Console.WriteLine("Finish get object from xml");
                if (_xmlReader != null)
                {
                    _xmlReader.Close();
                }
                if (_strReader != null)
                {
                    _strReader.Close();
                }
            }
            return _obj;
        }

        /// <summary>
        /// Format text to xml 
        /// Input value must be xml content
        /// </summary>
        /// <param name="xmlContent">Xml content</param>
        /// <returns></returns>
        public static String FormatXmlContent(String xmlContent)
        {
            try
            {
                XmlDocument document = new XmlDocument();
                document.Load(new StringReader(xmlContent));

                StringBuilder builder = new StringBuilder();
                using (XmlTextWriter writer = new XmlTextWriter(new StringWriter(builder)))
                {
                    writer.Formatting = Formatting.Indented;
                    document.Save(writer);
                }

                return builder.ToString();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                Console.WriteLine("Write file success");
            }
            return "";
        }
    }
}
