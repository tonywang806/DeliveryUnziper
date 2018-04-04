using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using SharpCompress.Archives.Zip;
using SharpCompress.Archives.Rar;

using SharpCompress.Readers;
using SharpCompress.Archives;
using SharpCompress.Common;
using System.IO;

namespace DeliveryUnziper
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("納品物の解凍を開始いたします。");
            Console.WriteLine("========================================");
            try
            {
                string CompressFileName = args[0];
#if DEBUG
                CompressFileName = @"C:\DeliveryUnziper\test\20180403_＜納品＞ECT_得意先向け在庫表発行（任意先）No_3496.rar";
#endif 
                FileInfo fi = new FileInfo(CompressFileName);
                switch (fi.Extension.ToLower())
                {
                    case ".zip":
                        ExtractFilesZip(fi);
                        break;
                    case ".rar":
                        ExtractFilesRar(fi);
                        break;
                    default:
                        throw new Exception("無効な圧縮ファイルです。ご確認のほど、お願いいたします。");
                }

                File.Delete(CompressFileName);
                Console.WriteLine("========================================");
                Console.WriteLine("納品物の解凍を正常終了しました。");
 #if DEBUG
                Console.ReadLine();
#endif
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.WriteLine("========================================");
                Console.WriteLine("納品物の解凍を異常終了しました。");
                Console.ReadLine();
            }

        }

        private static void ExtractFilesZip(FileInfo fi)
        {

            Console.WriteLine("DirectoryName:{0}", fi.DirectoryName);
            Console.WriteLine("FileName:{0}", fi.Name);
            Console.WriteLine("Extension:{0}", fi.Extension);


            string OutBasePath = fi.DirectoryName;

            ReaderOptions ro = new ReaderOptions();
            ro.ArchiveEncoding.Default = Encoding.GetEncoding(932);


            using (var archive = ZipArchive.Open(fi, ro))
            {
                //File毎で解凍
                string innerBaseFold = archive.Entries.FirstOrDefault().Key;

                foreach (var entry in archive.Entries)
                {
                    if (entry.Key == innerBaseFold)
                    {
                        continue;
                    }
                    string DestinationFileName = entry.Key.Replace(innerBaseFold, string.Empty);
                    DestinationFileName = Path.Combine(OutBasePath, DestinationFileName);

                    if (entry.IsDirectory)
                    {
                        Directory.CreateDirectory(DestinationFileName);
                    }
                    else
                    {

                        entry.WriteToFile(DestinationFileName, new ExtractionOptions()
                        {
                            Overwrite = true
                        });
                    }
                }
            }
        }
        private static void ExtractFilesRar(FileInfo fi)
        {
            Console.WriteLine("DirectoryName:{0}", fi.DirectoryName);
            Console.WriteLine("FileName:{0}", fi.Name);
            Console.WriteLine("Extension:{0}", fi.Extension);


            string OutBasePath = fi.DirectoryName;

            ReaderOptions ro = new ReaderOptions();
            ro.ArchiveEncoding.Default = Encoding.GetEncoding(932);

            using (var archive = RarArchive.Open(fi,ro))
            {
                //File毎で解凍

                //Rar専用ロジック①Entry List　Reverse
                IEnumerable<IArchiveEntry> entries = archive.Entries.Reverse();

                string innerBaseFold = entries.FirstOrDefault().Key;
                foreach (var entry in entries)
                {
                    if (entry.Key == innerBaseFold)
                    {
                        continue;
                    }

                    //Rar専用ロジック②Entry.Key
                    string DestinationFileName = entry.Key.Replace(String.Format("{0}\\",innerBaseFold), string.Empty);
                    DestinationFileName = Path.Combine(OutBasePath, DestinationFileName);

                    if (entry.IsDirectory)
                    {
                        Directory.CreateDirectory(DestinationFileName);
                    }
                    else
                    {

                        entry.WriteToFile(DestinationFileName, new ExtractionOptions()
                        {
                            Overwrite = true
                        });
                    }
                }
            }
        }

    }

}
