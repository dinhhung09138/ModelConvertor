using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace ModelConverter
{
    public class FileHelper
    {
        /// <summary>
        /// Write text file
        /// </summary>
        /// <param name="filePath">File path</param>
        /// <param name="content">Content</param>
        public static void WriteFile(string filePath, string content)
        {
            try
            {
                if (File.Exists(filePath))
                {
                    File.Delete(filePath);
                }
                File.WriteAllText(filePath, content, Encoding.UTF8);
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                Console.WriteLine("Write file success");
            }
        }

        public static void WriteFile(string filePath, string content, bool deleteOldFile)
        {
            try
            {
                if (File.Exists(filePath) && deleteOldFile)
                {
                    File.Delete(filePath);
                }
                File.WriteAllText(filePath, content, Encoding.UTF8);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                Console.WriteLine("Write file success");
            }
        }

        /// <summary>
        /// Read content from file path
        /// </summary>
        /// <param name="filePath">File path</param>
        /// <returns></returns>
        public static string ReadTextFile(string filePath)
        {
            try
            {
                using (StreamReader sr = new StreamReader(filePath))
                {
                    String line = sr.ReadToEnd();
                    return line.ToString();
                }
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
