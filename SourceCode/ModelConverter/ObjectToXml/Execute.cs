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

        public void Main()
        {
            string _filePath = Path.Combine(Environment.CurrentDirectory, "Employee.xml");

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

            //ReadAttribute(_filePath, "ID", "");
            //ReadAttribute(_filePath, "Details", "ProductID");
            //ReadAttribute(_filePath, "ProductID", "Node");
            //GetAttributeValueInSingleNodeFromXmlFile(_filePath, "/Employee/Details/DetailID/@ID");
            XMLContentModel classModel = new XMLContentModel()
            {
                RootName = "Classes",
                Comment = "",
                Nodes = new List<XMLNodeModel>()
                {
                    new XMLNodeModel()
                    {
                        Data = new XmlNodeData()
                        {
                            Name = "Class",
                            Value = ""
                        },
                        Properties = new List<XmlNodeData>()
                        {
                            new XmlNodeData()
                            {
                                Name = "ID",
                                Value = "10A1"
                            },
                            new XmlNodeData()
                            {
                                Name = "Name",
                                Value = "Class 10A1"
                            },
                            new XmlNodeData()
                            {
                                Name = "NumOfStudent",
                                Value = "100"
                            }
                        },
                        Childs = new List<XMLNodeModel>()
                        {
                            new XMLNodeModel()
                            {
                                Data = new XmlNodeData()
                                {
                                    Name = "Employees",
                                    Value =""
                                },
                                Childs = new List<XMLNodeModel>()
                                {
                                    new XMLNodeModel()
                                    {
                                        Data = new XmlNodeData()
                                        {
                                            Name = "Student",
                                            Value = ""
                                        },
                                        Properties = new List<XmlNodeData>()
                                        {
                                            new XmlNodeData()
                                            {
                                                Name = "Name",
                                                Value = "Hung"
                                            },
                                            new XmlNodeData()
                                            {
                                                Name = "Birthday",
                                                Value= DateTime.Now.ToString()
                                            }
                                        },
                                        Childs = new List<XMLNodeModel>()
                                        {
                                            new XMLNodeModel()
                                            {
                                                Data = new XmlNodeData()
                                                {
                                                    Name = "Address",
                                                    Value = "Go Vap"
                                                }
                                            },
                                            new XMLNodeModel()
                                            {
                                                Data = new XmlNodeData()
                                                {
                                                    Name = "Email",
                                                    Value = "Dinhhung@gmail.com"
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            };

            CreateXmlFile(Path.Combine(Environment.CurrentDirectory, "CreateEmployee.xml"), classModel);
            RemoveNode(Path.Combine(Environment.CurrentDirectory, "CreateEmployee.xml"), "/Classes/Class/Employees/Student/Email");
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

        public string ReadAttribute(string filePath, string nodeName, string attributeName)
        {
            try
            {
                XmlDocument _doc = new XmlDocument();
                _doc.Load(filePath);
                XmlNodeList _nodeList = _doc.GetElementsByTagName(nodeName);
                for (int i = 0; i < _nodeList.Count; i++)
                {
                    //Console.WriteLine(_nodeList[i].InnerText);
                    Console.WriteLine(_nodeList[i].InnerXml);
                    if (attributeName.Length > 0)
                    {
                        Console.WriteLine(_nodeList[i].Attributes[attributeName].Value);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                Console.WriteLine("Read attribute file success");
            }
            return "";
        }

        public string ReadAttribute(string filePath, string nodeName)
        {
            try
            {
                XmlDocument _doc = new XmlDocument();
                _doc.Load(filePath);
                XmlNodeList _nodeList = _doc.GetElementsByTagName(nodeName);
                var _tem = _doc.SelectSingleNode(nodeName);
                var _listItem = _doc.SelectNodes(nodeName);

                Console.WriteLine(_doc.SelectNodes(nodeName));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                Console.WriteLine("Read attribute file success");
            }
            return "";
        }

        /// <summary>
        /// Get value in single node
        /// </summary>
        /// <param name="filePath">Xml file path</param>
        /// <param name="nodePath">Node path. Ex:root/node/</param>
        /// <returns>String value, empty if node is not found or error</returns>
        private string GetValueInSingleNodeFromXmlFile(string filePath, string nodePath)
        {
            try
            {
                XmlDocument _doc = new XmlDocument();
                _doc.Load(filePath);
                var _item = _doc.SelectSingleNode(nodePath);
                if (_item != null)
                {
                    Console.WriteLine(_item.InnerText);
                    return _item.InnerText;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                Console.WriteLine("Read attribute file success");
            }
            return "";
        }

        /// <summary>
        /// Get value in single node
        /// </summary>
        /// <param name="xmlContent">Xml string</param>
        /// <param name="nodePath">Node path. Ex:root/node/</param>
        /// <returns>String value, empty if node is not found or error</returns>
        private string GetValueInSingleNodeFromXmlContent(string xmlContent, string nodePath)
        {
            try
            {
                XmlDocument _doc = new XmlDocument();
                _doc.LoadXml(xmlContent);
                var _item = _doc.SelectSingleNode(nodePath);
                if (_item != null)
                {
                    Console.WriteLine(_item.InnerText);
                    return _item.InnerText;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                Console.WriteLine("Read attribute file success");
            }
            return "";
        }

        /// <summary>
        /// Get XML content inside a single node
        /// </summary>
        /// <param name="filePath">Xml file path</param>
        /// <param name="nodePath">ode path. Ex:root/node/</param>
        /// <returns>Xml string content, empty if node is not found or error</returns>
        private string GetSingleNodeContentFromXmlFile(string filePath, string nodePath)
        {
            try
            {
                XmlDocument _doc = new XmlDocument();
                _doc.Load(filePath);
                var _item = _doc.SelectSingleNode(nodePath);
                if (_item != null)
                {
                    Console.WriteLine(_item.InnerXml);
                    return _item.InnerXml;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                Console.WriteLine("Read attribute file success");
            }
            return "";
        }

        /// <summary>
        /// Get XML content inside a single node
        /// </summary>
        /// <param name="xmlContent">Xml content</param>
        /// <param name="nodePath">ode path. Ex:root/node/</param>
        /// <returns>Xml string content, empty if node is not found or error</returns>
        private string GetSingleNodeContentFromXmlContent(string xmlContent, string nodePath)
        {
            try
            {
                XmlDocument _doc = new XmlDocument();
                _doc.LoadXml(xmlContent);
                var _item = _doc.SelectSingleNode(nodePath);
                if (_item != null)
                {
                    Console.WriteLine(_item.InnerXml);
                    return _item.InnerXml;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                Console.WriteLine("Read attribute file success");
            }
            return "";
        }

        /// <summary>
        /// Get attribute value in single node
        /// </summary>
        /// <param name="filePath">Xml file path</param>
        /// <param name="nodePath">Node path. Ex:root/node/@attribute</param>
        /// <returns>String value, empty if node is not found or error</returns>
        private string GetAttributeValueInSingleNodeFromXmlFile(string filePath, string nodePath)
        {
            try
            {
                XmlDocument _doc = new XmlDocument();
                _doc.Load(filePath);
                var _item = _doc.SelectSingleNode(nodePath);
                if (_item != null)
                {
                    Console.WriteLine(_item.InnerText);
                    return _item.InnerText;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                Console.WriteLine("Read attribute file success");
            }
            return "";
        }

        /// <summary>
        /// Get attribute value in single node
        /// </summary>
        /// <param name="xmlContent">Xml content</param>
        /// <param name="nodePath">Node path. Ex:root/node/@attribute</param>
        /// <returns>String value, empty if node is not found or error</returns>
        private string GetAttributeValueInSingleNodeFromXmlContent(string xmlContent, string nodePath)
        {
            try
            {
                XmlDocument _doc = new XmlDocument();
                _doc.LoadXml(xmlContent);
                var _item = _doc.SelectSingleNode(nodePath);
                if (_item != null)
                {
                    Console.WriteLine(_item.InnerText);
                    return _item.InnerText;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                Console.WriteLine("Read attribute file success");
            }
            return "";
        }

        public bool CreateXmlFile(string filePath, XMLContentModel node)
        {
            try
            {
                XmlWriterSettings _setting = new XmlWriterSettings();
                _setting.Indent = true;
                _setting.IndentChars = (" ");
                _setting.CloseOutput = true;
                _setting.OmitXmlDeclaration = false; //False: show xml version=1.0 encoding=utf-8, true: hidden
                using (XmlWriter writer = XmlWriter.Create(filePath, _setting))
                {
                    if (node.RootName.Length == 0)
                    {
                        return false;
                    }
                    if (node.Comment.Length > 0)
                    {
                        writer.WriteComment(node.Comment);
                    }
                    writer.WriteStartElement(node.RootName);
                    WriteNode(writer, node.Nodes);
                    writer.WriteEndElement();
                    //
                    writer.Flush();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            return false;
        }

        private bool WriteNode(XmlWriter writer, List<XMLNodeModel> nodes)
        {
            try
            {
                foreach (var n in nodes)
                {
                    if (n.Comment.Length > 0)
                    {
                        writer.WriteComment(n.Comment);
                    }
                    if (n.Data.Name.Length == 0)
                    {
                        continue;
                    }
                    if (n.Data.Value.Length > 0)
                    {
                        writer.WriteElementString(n.Data.Name, n.Data.Value);
                    }
                    else
                    {
                        writer.WriteStartElement(n.Data.Name);
                        if (n.Properties.Count > 0)
                        {
                            WriteAttribute(writer, n.Properties);
                        }
                        WriteNode(writer, n.Childs);
                        writer.WriteEndElement();
                    }
                }

            }
            catch (Exception ex)
            {

            }
            return true;
        }

        private bool WriteAttribute(XmlWriter writer, List<XmlNodeData> properties)
        {
            try
            {
                foreach (var p in properties)
                {
                    writer.WriteAttributeString(p.Name, p.Value);
                }
                return true;
            }
            catch (Exception ex)
            {

            }
            return false;
        }



        public bool RemoveNode(string filePath, string nodePath)
        {
            try
            {
                XMLNodeModel node = new XMLNodeModel()
                {
                    Data = new XmlNodeData()
                    {
                        Name = "Student",
                        Value = ""
                    },
                    Properties = new List<XmlNodeData>()
                    {
                        new XmlNodeData()
                        {
                            Name = "Name",
                            Value = "Phuc"
                        },
                        new XmlNodeData()
                        {
                            Name = "Birthday",
                            Value= DateTime.Now.AddYears(-20).ToString()
                        }
                    },
                    Childs = new List<XMLNodeModel>()
                    {
                        new XMLNodeModel()
                        {
                            Data = new XmlNodeData()
                            {
                                Name = "Address",
                                Value = "Phan thiet"
                            }
                        },
                        new XMLNodeModel()
                        {
                            Data = new XmlNodeData()
                            {
                                Name = "Email",
                                Value = "Dinhhung@gmail.com"
                            }
                        }
                    }
                };
                //
                XmlDocument _doc = new XmlDocument();
                _doc.Load(filePath);
                var _item = _doc.SelectSingleNode(nodePath);
                if(_item != null)
                {
                    _item.ParentNode.RemoveChild(_item);
                }
                _doc.Save(filePath);
            }
            catch (Exception ex)
            {

            }
            return false;
        }
    }

    public class XMLContentModel
    {
        /// <summary>
        /// Root name
        /// </summary>
        public string RootName { get; set; } = "";

        /// <summary>
        /// Comment
        /// </summary>
        public string Comment { get; set; } = "";

        /// <summary>
        /// List of node
        /// </summary>
        public List<XMLNodeModel> Nodes { get; set; } = new List<XMLNodeModel>();
    }

    /// <summary>
    /// Xml node data model
    /// </summary>
    public class XMLNodeModel
    {
        /// <summary>
        /// Data of node
        /// </summary>
        public XmlNodeData Data { get; set; } = new XmlNodeData();

        /// <summary>
        /// Comment
        /// </summary>
        public string Comment { get; set; } = "";

        /// <summary>
        /// Child nodes
        /// </summary>
        public List<XMLNodeModel> Childs { get; set; } = new List<XMLNodeModel>();

        /// <summary>
        /// Properties of node
        /// </summary>
        public List<XmlNodeData> Properties { get; set; } = new List<XmlNodeData>();


    }

    /// <summary>
    /// Data of 
    /// </summary>
    public class XmlNodeData
    {
        /// <summary>
        /// Node name
        /// </summary>
        public string Name { get; set; } = "";

        /// <summary>
        /// Node value. If has value, Node can not write child node inside
        /// </summary>
        public string Value { get; set; } = "";
    }
}
