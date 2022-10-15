using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using ExcelDataReader;
using System.Data.SqlClient;


namespace ReadExcelFile
{
    public class DataReaderClass
    {
        public static string DataReader(string filePath)
        {
            string sonuc = "NOK";
            try
            {
                List<DateTimeOffset> tarihList = new List<DateTimeOffset>();
                List<string> personelList = new List<string>();
                List<string> vardiyaList = new List<string>();

                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                DataTable dt = new DataTable();
              //  string filePath = @"/Users/alperalanyali/Desktop/"+fileName+".xlsx";           
                FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.Read);
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var rs = reader.AsDataSet();
                    dt = rs.Tables[0];
                    Console.WriteLine(dt.Rows[2][0]);
                    Console.WriteLine(dt.Rows[3][0]);
                    for (int i = 2; i < dt.Rows.Count; i++)
                    {
                        //Console.WriteLine(dt.Rows[i][0]);
                        for (int j = 0; j < dt.Columns.Count; j++)
                        {
                       
                            /* if (j % 2 == 1){
                                 //Personel bilgisi
                                  Console.WriteLine("Tekler:" + dt.Rows[i][j]);
                                 personelList.Add(dt.Rows[i][j].ToString());
                             }else
                             {
                                if(j != 0)
                                 {
                                     //Vardiya Bilgisi
                                     Console.WriteLine("Çiftler:" + dt.Rows[i][j]);
                                     vardiyaList.Add(dt.Rows[i][j].ToString());
                                 } else
                                 {
                                     //Tarih bilgisi
                                     DateTimeOffset tarih = new DateTimeOffset();
                                     DateTimeOffset.TryParse(dt.Rows[i][j].ToString(), out tarih);
                                     tarihList.Add(tarih);
                                 }
                             }*/

                        }
                        
                      /* if(i == 1)
                       {
                            for (int j = 0; j <= dt.Columns.Count; j++)
                            {

                                Console.WriteLine(dt.Rows[i][j]+"\t");

                            }
                       }*/
                        
                 
                      /*  if (i > 1)
                        {
                            int k = 0;
                            for (int j = 0; j <= dt.Columns.Count; j++)
                            {
                                
                              Console.WriteLine(dt.Rows[i][j+1]);

                                
                            }

                        }*/
                        Console.WriteLine("------");
                    }
                    
                }
            }
            catch(Exception ex)
            {
                sonuc += " " + ex.Message;
            }
            return sonuc;
        }


        public static List<Guid> GetPersonelId(List<string> personelAdi)
        {
            List<Guid> personelId =new List<Guid>();

            try
            {
                for (int i = 0; i < personelAdi.Count; i++)
                {
                    string connectionString = "";
                    string q = "select * from Personel where select * from Personel where AdiSoyadi like '%" + personelAdi[i].ToString() + "%'";
                    SqlDataAdapter da = new SqlDataAdapter(q, connectionString);
                    DataTable dt = new DataTable();
                    da.Fill(dt);
                    Guid personel = new Guid();
                    Guid.TryParse(dt.Rows[0]["Id"].ToString(), out personel);
                    personelId.Add(personel);
                }
            }
            catch(Exception ex)
            {

            }

            return personelId;
        }
        public static string InsertIntoPersonelVardiyaAtama(List<DateTimeOffset> tarihler,List<string> personeller,List<string> vardiyalar)
        {
            string sonuc = "NOK";
            try
            {
                for (int i = 0; i < tarihler.Count; i++)
                {
                    string connectionString = "";
                    SqlConnection connection = new SqlConnection(connectionString);
                    connection.Open();
                    string insertQuery = " INSERT INTO PersonelVardiyaAtama (Tarih, Id, MusteriId, PersonelId, VardiyaId) VALUES (@Tarih, @Id, @MusteriId, @PersonelId, @VardiyaId)";
                    SqlCommand cmd = new SqlCommand();
                    cmd.Connection = connection;
                    cmd.Parameters.AddWithValue("@Tarih", tarihler[i]);
                    cmd.Parameters.AddWithValue("@Id", Guid.NewGuid());
                    cmd.Parameters.AddWithValue("@MusteriId", Guid.NewGuid());

                    connection.Close();
                    
                }

            }catch(Exception ex)
            {
                sonuc += " " + ex.Message;
                if(ex.InnerException != null)
                {
                    sonuc += " " + ex.InnerException.Message;
                }
            }
            return sonuc;
        }
    }
}
