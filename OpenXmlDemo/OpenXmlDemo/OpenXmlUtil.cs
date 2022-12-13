using System.Collections.Generic;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using AA = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace OpenXmlDemo
{
    public static class OpenXmlUtil
    {
        /// <summary>
        /// 按书签替换图片
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="picPath"></param>
        /// <param name="bm"></param>
        /// <param name="x">宽度厘米</param>
        /// <param name="y">高度厘米</param>
        /// <param name="type"></param>
        public static void ReplaceBMPicture(string filePath, string picPath, string bm) {
            RemoveBookMarkContent(filePath, bm);
            InsertBMPicture(filePath, picPath, bm);
        }

        /// <summary>
        /// 按书签替换图片
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="picPath"></param>
        /// <param name="bm"></param>
        /// <param name="x">宽度厘米</param>
        /// <param name="y">高度厘米</param>
        /// <param name="type"></param>
        public static void ReplaceBMPicture(string filePath, string picPath, string bm, long x, ImagePartType type) {
            RemoveBookMarkContent(filePath, bm);
            InsertBMPicture(filePath, picPath, bm, x, type);
        }
        /// <summary>
        /// 按书签替换图片
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="picPath"></param>
        /// <param name="bm"></param>
        /// <param name="x">宽度厘米</param>
        /// <param name="y">高度厘米</param>
        /// <param name="type"></param>
        public static void ReplaceBMPicture(string filePath, string picPath, string bm, long x, long y, ImagePartType type) {
            RemoveBookMarkContent(filePath, bm);
            InsertBMPicture(filePath, picPath, bm, x, y, type);
        }
        /// <summary>
        /// 按书签插入图片
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="picPath"></param>
        /// <param name="bm"></param>
        /// <param name="x">宽度厘米</param>
        /// <param name="y">高度厘米</param>
        /// <param name="type"></param>
        public static void InsertBMPicture(string filePath, string picPath, string bm, long x, ImagePartType type) {
            long y = 0;
            using (System.Drawing.Bitmap objPic = new System.Drawing.Bitmap(picPath)) {
                y = (x * objPic.Height) / objPic.Width;
            }
            InsertBMPicture(filePath, picPath, bm, x, y, type);
        }

        /// <summary>
        /// 按书签插入图片
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="picPath"></param>
        /// <param name="bm"></param>
        /// <param name="x">宽度厘米</param>
        /// <param name="y">高度厘米</param>
        /// <param name="type"></param>
        public static void InsertBMPicture(string filePath, string picPath, string bm, long x, long y, ImagePartType type) {
            using (WordprocessingDocument doc =
               WordprocessingDocument.Open(filePath, true)) {
                MainDocumentPart mainPart = doc.MainDocumentPart;

                BookmarkStart bmStart = findBookMarkStart(doc, bm);
                if (bmStart == null) {
                    return;
                }

                ImagePart imagePart = mainPart.AddImagePart(type);

                using (FileStream stream = new FileStream(picPath, FileMode.Open)) {
                    imagePart.FeedData(stream);
                }
                long cx = 360000L * x;//360000L = 1厘米
                long cy = 360000L * y;
                Run r = AddImageToBody(doc, mainPart.GetIdOfPart(imagePart), cx, cy);
                bmStart.Parent.InsertAfter<Run>(r, bmStart);
                mainPart.Document.Save();
            }
        }
        /// <summary>
        /// 按书签插入图片。默认15厘米，JPG
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="picPath"></param>
        /// <param name="bm"></param>
        public static void InsertBMPicture(string filePath, string picPath, string bm) {
            InsertBMPicture(filePath, picPath, bm, 15, 15, ImagePartType.Jpeg);
        }
        /// <summary>
        /// 查找书签
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="bmName"></param>
        /// <returns></returns>
        private static BookmarkStart findBookMarkStart(WordprocessingDocument doc, string bmName) {
            foreach (var footer in doc.MainDocumentPart.FooterParts) {
                foreach (var inst in footer.Footer.Descendants<BookmarkStart>()) {
                    if (inst.Name == bmName) {
                        return inst;
                    }
                }
            }

            foreach (var header in doc.MainDocumentPart.HeaderParts) {
                foreach (var inst in header.Header.Descendants<BookmarkStart>()) {
                    if (inst.Name == bmName) {
                        return inst;
                    }
                }
            }
            foreach (var inst in doc.MainDocumentPart.RootElement.Descendants<BookmarkStart>()) {
                if (inst is BookmarkStart) {
                    if (inst.Name == bmName) {
                        return inst;
                    }
                }
            }

            return null;
        }

        /// <summary>
        /// 查找书签
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="bmName"></param>
        /// <returns></returns>
        private static List<BookmarkStart> findAllBookMarkStart(WordprocessingDocument doc) {
            List<BookmarkStart> ret = new List<BookmarkStart>();
            foreach (var footer in doc.MainDocumentPart.FooterParts) {
                ret.AddRange(footer.Footer.Descendants<BookmarkStart>());

            }
            foreach (var header in doc.MainDocumentPart.HeaderParts) {
                ret.AddRange(header.Header.Descendants<BookmarkStart>());
            }
            ret.AddRange(doc.MainDocumentPart.RootElement.Descendants<BookmarkStart>());
            return ret;
        }
        /// <summary>
        /// 查找书签
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="bmName"></param>
        /// <returns></returns>
        private static List<BookmarkEnd> findAllBookMarkEnd(WordprocessingDocument doc) {
            List<BookmarkEnd> ret = new List<BookmarkEnd>();
            foreach (var footer in doc.MainDocumentPart.FooterParts) {
                ret.AddRange(footer.Footer.Descendants<BookmarkEnd>());

            }
            foreach (var header in doc.MainDocumentPart.HeaderParts) {
                ret.AddRange(header.Header.Descendants<BookmarkEnd>());
            }
            ret.AddRange(doc.MainDocumentPart.RootElement.Descendants<BookmarkEnd>());
            return ret;
        }


        /// <summary>
        /// 查找书签END
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="bmName"></param>
        /// <returns></returns>
        private static BookmarkEnd findBookMarkEnd(WordprocessingDocument doc, string id) {
            foreach (var footer in doc.MainDocumentPart.FooterParts) {
                foreach (var inst in footer.Footer.Descendants<BookmarkEnd>()) {
                    if (inst.Id == id) {
                        return inst;
                    }
                }
            }

            foreach (var header in doc.MainDocumentPart.HeaderParts) {
                foreach (var inst in header.Header.Descendants<BookmarkEnd>()) {
                    if (inst.Id == id) {
                        return inst;
                    }
                }
            }
            foreach (var inst in doc.MainDocumentPart.RootElement.Descendants<BookmarkEnd>()) {
                if (inst.Id == id) {
                    return inst;
                }

            }

            return null;
        }

        private static Run AddImageToBody(WordprocessingDocument wordDoc, string relationshipId, long cx, long cy) {
            return new Run(new Drawing(
                     new DW.Inline(
                         new DW.Extent() { Cx = cx, Cy = cy },
                         new DW.EffectExtent() {
                             LeftEdge = 0L,
                             TopEdge = 0L,
                             RightEdge = 0L,
                             BottomEdge = 0L
                         },
                         new DW.DocProperties() {
                             Id = (UInt32Value)1U,
                             Name = "Picture 1"
                         },
                         new DW.NonVisualGraphicFrameDrawingProperties(
                             new AA.GraphicFrameLocks() { NoChangeAspect = true }),
                         new AA.Graphic(
                             new AA.GraphicData(
                                 new PIC.Picture(
                                     new PIC.NonVisualPictureProperties(
                                         new PIC.NonVisualDrawingProperties() {
                                             Id = (UInt32Value)0U,
                                             Name = "New Bitmap Image.jpg"
                                         },
                                         new PIC.NonVisualPictureDrawingProperties()),
                                     new PIC.BlipFill(
                                         new AA.Blip(
                                             new AA.BlipExtensionList(
                                                 new AA.BlipExtension() {
                                                     Uri =
                                                       "{28A0092B-C50C-407E-A947-70E740481C1C}"
                                                 })
                                         ) {
                                             Embed = relationshipId,
                                             CompressionState =
                                             AA.BlipCompressionValues.Print
                                         },
                                         new AA.Stretch(
                                             new AA.FillRectangle())),
                                     new PIC.ShapeProperties(
                                         new AA.Transform2D(
                                             new AA.Offset() { X = 0L, Y = 0L },
                                             new AA.Extents() { Cx = 990000L, Cy = 792000L }),
                                         new AA.PresetGeometry(
                                             new AA.AdjustValueList()
                                         ) { Preset = AA.ShapeTypeValues.Rectangle }))
                             ) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                     ) {
                         DistanceFromTop = (UInt32Value)0U,
                         DistanceFromBottom = (UInt32Value)0U,
                         DistanceFromLeft = (UInt32Value)0U,
                         DistanceFromRight = (UInt32Value)0U,
                         EditId = "50D07946"
                     }));

            // Append the reference to body, the element should be in a Run.
            //  wordDoc.MainDocumentPart.Document.Body.AppendChild(new Paragraph(new Run(element)));
        }

        public static void DeleteRange(string filePath, string stringStart, string stringStop, int way) {
            using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, true)) {
                List<OpenXmlElement> list = RangeFind(doc, stringStart, stringStop, way);
                foreach (var inst in list) {
                    inst.Remove();
                }
            }
        }
        /// <summary>
        /// 1 标记1结束到标记2开始;2 标记1结束到标记2结束;3 标记1开始到标记2结束; 4 标记1开始到标记2开始;
        /// trimhuiche 如果为true，则考虑回车；否则不考虑回车。
        /// chzhao@wisdombud.com
        /// </summary>
        /// <param name="stringStart"></param>
        /// <param name="stringStop"></param>
        /// <param name="way"></param>
        public static List<OpenXmlElement> RangeFind(WordprocessingDocument doc, string stringStart, string stringStop, int way) {
            List<OpenXmlElement> ret = new List<OpenXmlElement>();
            bool add = false;
            foreach (var inst in doc.MainDocumentPart.Document.Body.Elements()) {

                if (way == 1) {
                    if (inst.InnerText.Contains(stringStop)) {
                        add = false;
                    }
                    if (add) {
                        ret.Add(inst.CloneNode(true));
                    }
                    if (inst.InnerText == stringStart) {
                        add = true;
                    }
                }
                else if (way == 2) {

                    if (add) {
                        ret.Add(inst.CloneNode(true));
                    }
                    if (inst.InnerText == stringStart) {
                        add = true;
                    }
                    if (inst.InnerText.Contains(stringStop)) {
                        add = false;
                    }
                }
                else if (way == 3) {
                    if (inst.InnerText == stringStart) {
                        add = true;
                    }
                    if (add) {
                        ret.Add(inst.CloneNode(true));
                    }

                    if (inst.InnerText.Contains(stringStop)) {
                        add = false;
                    }
                }
                else if (way == 4) {
                    if (inst.InnerText == stringStart) {
                        add = true;
                    }
                    if (inst.InnerText.Contains(stringStop)) {
                        add = false;
                    }
                    if (add) {
                        ret.Add(inst.CloneNode(true));
                    }

                }



            }
            return ret;
        }


        /// <summary>
        /// 修改书签
        /// </summary>
        /// <param name="filePath">word文档</param>
        /// <param name="bmName">书签名字</param>
        /// <param name="text">替换的文本</param>
        public static void ModifyBM(string filePath, string bmName, string text) {
            using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, true)) {
                BookmarkStart bmStart = findBookMarkStart(doc, bmName);

                Run bookmarkText = bmStart.NextSibling<Run>();
                if (bookmarkText != null) {
                    Text t = bookmarkText.GetFirstChild<Text>();
                    if (t != null) {
                        t.Text = text;
                    }
                }
            }
        }

        /// <summary>
        /// 替换书签内容
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="bmName"></param>
        /// <param name="text"></param>
        public static void InsertIntoBookmark(string filePath, string bmName, string text) {
            using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, true)) {
                BookmarkStart bmStart = findBookMarkStart(doc, bmName);
                BookmarkEnd bmEnd = findBookMarkEnd(doc, bmStart.Id);
                if (bmStart == null) {
                    return;
                }
                var run = bmStart.NextSibling();
                while (run != null && !(run is BookmarkEnd)) {
                    OpenXmlElement nextElem = run.NextSibling();
                    run.Remove();
                    run = nextElem;
                }
                bmStart.Parent.InsertAfter<Run>(new Run(new Text(text)), bmStart);
                doc.Save();
            }
        }

        /// <summary>
        /// 删除书签内容
        /// </summary>
        /// <param name="bookmark"></param>
        public static void RemoveBookMarkContent(string filePath, string bmName) {
            using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, true)) {
                BookmarkStart bmStart = findBookMarkStart(doc, bmName);
                BookmarkEnd bmEnd = findBookMarkEnd(doc, bmStart.Id);
                while (true) {
                    var run = bmStart.NextSibling();
                    if (run == null) {
                        break;
                    }
                    if (run is BookmarkEnd && (BookmarkEnd)run == bmEnd) {
                        break;
                    }

                    run.Remove();
                }

            }
        }
        /// <summary>
        /// 重命名书签，在书签前面加前缀
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="prefix">前缀</param>
        public static void RenameBookMark(string filePath, string prefix) {
            using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, true)) {
                foreach (var inst in findAllBookMarkStart(doc)) {
                    inst.Name = prefix + inst.Name;
                }
            }
        }

        /// <summary>
        /// 重命名书签
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="oldName"></param>
        /// <param name="newName"></param>
        public static void RenameBookMark(string filePath, string oldName, string newName) {
            using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, true)) {
                var bm = findBookMarkStart(doc, oldName);
                bm.Name = newName;
            }
        }

        /// <summary>
        /// 删除书签
        /// </summary>
        /// <param name="bookmark"></param>
        public static void RemoveBookMark(string filePath, string bmName) {
            using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, true)) {
                var bmStart = findBookMarkStart(doc, bmName);
                if (bmStart == null) {
                    return;
                }
                var bmEnd = findBookMarkEnd(doc, bmStart.Id);
                bmStart.Remove();
                bmEnd.Remove();

            }
        }
        /// <summary>
        /// 合并文档
        /// </summary>
        /// <param name="finalFile"></param>
        /// <param name="files"></param>
        public static void Combine(string finalFile, List<string> files) {
            if (files.Count < 2) {
                return;
            }
            File.Copy(files[0], finalFile, true);
            using (WordprocessingDocument doc = WordprocessingDocument.Open(finalFile, true)) {
                Body b = doc.MainDocumentPart.Document.Body;
                for (int i = 1; i < files.Count; i++) {
                    using (WordprocessingDocument doc1 = WordprocessingDocument.Open(files[i], true)) {
                        foreach (var inst in doc1.MainDocumentPart.Document.Body.Elements()) {
                            b.Append(inst.CloneNode(true));
                        }
                    }
                }
            }
        }
    }
}