using DocumentFormat.OpenXml.Packaging;
using ImageMagick;
using MediaToolkit;
using MediaToolkit.Model;
using System;
using System.IO;
using System.Linq;

namespace SimpleMediaCompress
{
    public class MediaCompress
    {
        private readonly string _tempFileLocation = @"";
        public MediaCompress(string tempFileLocation)
        {            
            if (!Directory.Exists(tempFileLocation))
            {
                throw new Exception(tempFileLocation + " folder root must be created.");
            }
            _tempFileLocation = tempFileLocation;
        }

        public byte[] CompressImage(byte[] inputBytes)
        {
            try
            {
                using (MagickImage image = new MagickImage(inputBytes))
                {
                    image.Quality = 30;
                    image.Format = MagickFormat.Jpeg;
                    image.Strip();
                    return image.ToByteArray();
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Error on image compress = " + ex.Message);
            }
            
        }

        public byte[] CompressVideo(byte[] inputVideoBytes)
        {
            string tempInputPath = Path.Combine(_tempFileLocation, Guid.NewGuid().ToString() + ".mp4");
            string tempOutputPath = Path.Combine(_tempFileLocation, Guid.NewGuid().ToString() + "_cmp.mp4");

            try
            {
                File.WriteAllBytes(tempInputPath, inputVideoBytes);

                MediaFile inputFile = new MediaFile { Filename = tempInputPath };
                MediaFile outputFile = new MediaFile { Filename = tempOutputPath };

                using (var engine = new Engine())
                {
                    engine.GetMetadata(inputFile);
                    var arguments = $"-i \"{tempInputPath}\" -vcodec libx264 -crf 28 -preset slow -acodec aac -b:a 128k \"{tempOutputPath}\"";
                    engine.CustomCommand(arguments);
                }

                return File.ReadAllBytes(tempOutputPath);
            }
            catch (Exception ex)
            {
                throw new Exception("Error on video compress = " + ex.Message);
            }
            finally
            {
                if (File.Exists(tempInputPath))
                    File.Delete(tempInputPath);

                if (File.Exists(tempOutputPath))
                    File.Delete(tempOutputPath);
            }
        }

        public byte[] CompressPptx(byte[] pptxBytes)
        {
            string tempPath = Path.Combine(_tempFileLocation, Guid.NewGuid().ToString() + ".pptx");
            try
            {
                File.WriteAllBytes(tempPath, pptxBytes);

                using (PresentationDocument presentation = PresentationDocument.Open(tempPath, true))
                {
                    foreach (var slidePart in presentation.PresentationPart.SlideParts)
                    {
                        var images = slidePart.GetPartsOfType<ImagePart>().ToList();

                        foreach (var imagePart in images)
                        {
                            using (var stream = imagePart.GetStream())
                            {
                                using (var ms = new MemoryStream())
                                {
                                    if (imagePart.ContentType == "image/jpeg" || imagePart.ContentType == "image/png"
                                        || imagePart.ContentType == "image/jpg")
                                    {
                                        stream.CopyTo(ms);
                                        var originalBytes = ms.ToArray();
                                        var compressedBytes = CompressImage(originalBytes);
                                        using (var writeStream = imagePart.GetStream(FileMode.Create, FileAccess.Write))
                                        {
                                            writeStream.Write(compressedBytes, 0, compressedBytes.Length);
                                        }
                                    }                                        
                                }
                            }                            
                        }
                    }
                }
                return File.ReadAllBytes(tempPath);
            }
            catch (Exception ex)
            {
                throw new Exception("Error on pptx compress = " + ex.Message);
            }
            finally
            {
                if (File.Exists(tempPath))
                    File.Delete(tempPath);
            }
        }

    }
}
