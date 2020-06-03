using System;
using System.Data;
using System.Threading;
//using Microsoft.AnalysisServices;
//using Microsoft.AnalysisServices.Xmla;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Globalization;
using System.Data.OleDb;
using System.Xml;
namespace Test
{

    class Program
    {

        public static void Main(string[] args)
        {

            string decode = "Зьйв" + "даез" + "клип" + "омДЕ" +//128
                          "ЙжЖф" + "цтыщ" + "яЦЬш" + " ШЧ " + //128+16
                          "бнуъ" + "сСЄє" + "ї   " + "    " + //128+32
                          "    " + " БВА" + "    " + "  Ґ " + //128+48
                          "    " + "  гГ" + "    " + "    " +
                          "рРКЛ" + "И НО" + "П   " + "  М " +
                          "УЯФТ" + "хХ ю" + "ЮЪЫЩ" + "эЭЇґ" +
                          "-   " + "  ч " + "   №" + "іІ  ";

            var arr= System.Text.Encoding.GetEncoding(1251).GetBytes(decode);
            byte[] ResB = new byte[128];


            for (byte i = 0; i < 128; i++)
            {
                if(arr[i]>128)
                   ResB[arr[i]-128] =(byte)(i+(byte)128);
            }
            ByteArrayToFile(@"d:\to_sewoo_lk.map", ResB);

            return;


            var ll = decode.Length;


            List<byte> list = new List<byte>();

            list.AddRange(Encoding.ASCII.GetBytes(@"! 0 200 200 100 1
TONE 200
SPEED 5
PAGE-WIDTH  500
"
));
            //АБВГДЕЖЗИЙКЛМНОП РСТУФХЦЧШЩЪЫЬЭЮЯ абвгдежзийклмноп рстуфхцчшщъыьэюя іІїЇєЄҐґ
            string Test = "П";
            list.AddRange(Encoding.ASCII.GetBytes("SCALE-TEXT ANB4.tfd  36 36 0 0 "));

            /*            for (byte i = 0; i < 64; i++)
                        { 
                            //var r = decode.IndexOf(Ch);
                        list.Add((byte)(192 +i ));
                    }*/

            
                for (int i = 0; i < Test.Length; i++)
                {
                    var Ch = Test.Substring(i, 1);
                    if (Ch.ToCharArray()[0] > (char)127)
                    {
                        var r = decode.IndexOf(Ch);
                        list.Add((byte)(128 + r));
                    }
                    else
                        list.Add((byte) Ch.ToCharArray()[0]);
                }


            /*//Bild page
                        int j = 0;
                            int l = 0;
                            for (byte i = 188; i < 255; i++)
                            {


                                if (i % 4 == 0)
                                {
                                    list.Add(13);
                                    list.Add(10);
                                    list.AddRange(Encoding.ASCII.GetBytes("SCALE-TEXT ANB4.tfd  36 36 0 " + ((l++) * 40).ToString() + " "));
                                }
                list.AddRange(Encoding.ASCII.GetBytes(" "+ i.ToString()+ " "));
                list.Add(i);
                            }*/


            //list.Add(254);
            //list.Add(255);


            /*            int j = 0;
                        int l = 0;
                        for (byte i = 128; i < 255; i++)
                        {

                            if (i % 4 == 0)
                                list.Add((byte)'1');
                            if (i % 16 == 0)
                            {
                                list.Add(13);
                                list.Add(10);
                                list.AddRange(Encoding.ASCII.GetBytes("SCALE-TEXT ANB4.tfd  36 36 0 " + ((l++) * 40).ToString() + " 1"));
                            }
                            list.Add(i);
                        }*/


            list.AddRange(Encoding.ASCII.GetBytes(@"
PRINT
"));

            ByteArrayToFile(@"d:\11.txt", list.ToArray());
            return;
        }
        

        public static bool ByteArrayToFile(string fileName, byte[] byteArray)
        {
            try
            {
                using (var fs = new FileStream(fileName, FileMode.Create, FileAccess.Write))
                {
                    fs.Write(byteArray, 0, byteArray.Length);
                    return true;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception caught in process: {0}", ex);
                return false;
            }
        }
    }
}