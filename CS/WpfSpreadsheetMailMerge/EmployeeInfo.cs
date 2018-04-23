using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Xml.Serialization;

namespace WpfSpreadsheetMailMerge
{
    public static class EmployeesInfo
    {
        public static IEnumerable<EmployeeInfo> GetData()
        {
            return DataSourceHelper.GetDataSouresFromXml<EmployeeInfo>("WpfSpreadsheetMailMerge.data.xml", "ArrayOfEmployeeInfo");
        }
    }

    public static class DataSourceHelper
    {
        public static List<T> GetDataSouresFromXml<T>(string resName, string attribute) where T : class
        {
            using (Stream stream = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream(resName))
            {
                XmlSerializer s = new XmlSerializer(typeof(List<T>), new XmlRootAttribute(attribute));
                return (List<T>)s.Deserialize(stream);
            }
        }
    }

    public class EmployeeInfo
    {
        public string LastName { get; set; }
        public string FirstName { get; set; }
        public string Title { get; set; }
        public string TitleOfCourtesy { get; set; }
        public DateTime BirthDate { get; set; }
        public DateTime HireDate { get; set; }
        public string Address { get; set; }
        public string City { get; set; }
        public string Region { get; set; }
        public string PostalCode { get; set; }
        public string Country { get; set; }
        public string HomePhone { get; set; }
        public string Extension { get; set; }
        [XmlIgnoreAttribute()]
        public Bitmap Photo { get; set; }
        [XmlElementAttribute("Photo")]
        public byte[] PictureByteArray
        {
            get
            {
                if (Photo != null)
                {
                    TypeConverter BitmapConverter =
                         TypeDescriptor.GetConverter(Photo.GetType());
                    return (byte[])
                         BitmapConverter.ConvertTo(Photo, typeof(byte[]));
                }
                else
                    return null;
            }
            set
            {
                if (value != null)
                    Photo = new Bitmap(new MemoryStream(value));
                else
                    Photo = null;
            }
        }
        public string Notes { get; set; }
    }
}
