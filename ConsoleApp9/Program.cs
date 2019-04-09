using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media.Imaging;
namespace ConsoleApp9
{
    class Program
    {
        static void Main(string[] args)
        {
            DataTable dt = new DataTable();
            string con = @"Provider = Microsoft.ACE.OLEDB.12.0; Data Source = C:\Santhosh\tifFileDetails.xlsx; Extended Properties = 'Excel 12.0 Xml;HDR=YES;'";
            using (OleDbConnection connection = new OleDbConnection(con))
            {
               
                OleDbCommand command = new OleDbCommand("select * from [Sheet1$]", connection);
                using (OleDbCommand comm = new OleDbCommand())
                {
                    comm.CommandText = "select * from [Sheet1$]";
                    comm.Connection = connection;
                    using (OleDbDataAdapter da = new OleDbDataAdapter())
                    {
                        da.SelectCommand = comm;
                        da.Fill(dt);
                       
                    }

                }
            }
            List<int> vs = new List<int>();
            foreach (DataRow item in dt.AsEnumerable())
            {
                string[] files = Directory.GetFiles(item[1].ToString());
                foreach (var file in files)
                {
                    if (file.EndsWith(".tif"))
                    {
                        Stream imageStreamSource = new FileStream(file, FileMode.Open, FileAccess.Read, FileShare.Read);
                        TiffBitmapDecoder decoder = new TiffBitmapDecoder(imageStreamSource, BitmapCreateOptions.PreservePixelFormat, BitmapCacheOption.Default);
                        vs.Add(decoder.Frames.Count);
                        imageStreamSource.Close();
                    }
                }
               
            }

           
        }
    }
}
