using System;
using System.Configuration;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using ModelConverter.Model;
using ModelConverter.Excel;
using ModelConverter.ObjectToXml;
using System.Collections.Generic;
using NReco.VideoConverter;
using System.Drawing;

namespace ModelConverter
{
    class Program
    {
        static void Main(string[] args)
        {
            GetThumbnailVideo(Path.Combine(Directory.GetCurrentDirectory(), "video.mp4"), "video.mp4");
            return;
            ShowHelp();
            Console.WriteLine("Input value");
            var line = Console.ReadLine();
            switch (line)
            {
                case "1":
                    ExportModel model = new ExportModel();
                    model.Execute();
                    break;
                case "2":
                    ExportExcel excel = new ExportExcel();
                    excel.Execute();
                    break;
                case "3":
                    Execute exe = new Execute();
                    exe.Main();
                    break;
                case "0":
                    ShowHelp();
                    break;
                case "e":
                    Environment.Exit(0);
                    break;
                default:
                    Console.WriteLine("Not found");
                    break;
            }
            
        }

        private static void ShowHelp()
        {
            Console.WriteLine("---------------");
            Console.WriteLine("1. Export model");
            Console.WriteLine("2. Export excel");
            Console.WriteLine("3. ");
            Console.WriteLine("0: show help content");
            Console.WriteLine("e: exists");
            Console.WriteLine("");
            Console.WriteLine("---------------");
        }

        private static void GetThumbnailVideo(string videoPath, string videoName)
        {
            var videoFramesFolder = "VideoFrames";
            var rName = Guid.NewGuid() + ".png";
            string firstFrameFolderPhysical = Path.Combine(Directory.GetCurrentDirectory(), videoFramesFolder);

            if (Directory.Exists(firstFrameFolderPhysical))
            {
                Directory.Delete(firstFrameFolderPhysical, true);
            }
            Directory.CreateDirectory(firstFrameFolderPhysical);

            //string firstFrameFolderWeb = SiteConfigurationSetting.GetConfiguration("WebPath") +
            //    SiteConfigurationSetting.GetConfiguration("UserTempProjectPath") + "/" + videoFramesFolder + "/" + currentUserId + "/" + objTempModel.Id + "/" + cardFolderName + "/";

            var ffMpeg = new FFMpegConverter();
            //copy file to local folder and get thumbnail
            var rFolderName = videoName.Replace(".mp4", "");
            var webApiRoot = Path.Combine(Directory.GetCurrentDirectory(), rFolderName);
            if (Directory.Exists(webApiRoot))
            {
                Directory.Delete(webApiRoot, true);
            }
            Directory.CreateDirectory(webApiRoot);

            //var destvideoPhysical = webApiRoot + "\\" + videoFileName;
            //File.Copy(item, destvideoPhysical, true);
            var thumbnailPath = webApiRoot + "\\" + rName;
            ffMpeg.GetVideoThumbnail(videoPath, thumbnailPath);

            //resize
            Bitmap img = (Bitmap)Bitmap.FromFile(thumbnailPath);
            var image_p = ResizeImage(img, Convert.ToInt32(1200), Convert.ToInt32(900));
            image_p.Save(firstFrameFolderPhysical + "\\" + rName);
            image_p.Dispose();
            img.Dispose();

            //copy to gmacdn
            //File.Copy(thumbnailPath, firstFrameFolderPhysical + "\\" + rName);
            //firstFrameFolderWeb = firstFrameFolderWeb + "/" + rName;
            //delete Temp folder
            //Directory.Delete(webApiRoot, true);
        }

        private static Bitmap ResizeImage(Bitmap img, int width, int height)
        {
            Image.GetThumbnailImageAbort callback = new Image.GetThumbnailImageAbort(GetThumbAbort);
            return (Bitmap)img.GetThumbnailImage(width, height, callback, System.IntPtr.Zero);

        }

        public static bool GetThumbAbort()
        {
            return false;
        }

    }
    
    
}
