using Illustrator;
using NetTopologySuite.Geometries;
using PdfSharp.Drawing;
using PdfSharp.Fonts;
using PdfSharp.Pdf;
using System.Collections;

namespace o_w1
{
    
    class ColorIndex
    {
        private static int _globalIndex = 1;
        private int _index;
        public int index { get { return _index; } }
        public RGBColor color;
        public int Red { get { return (int)color.Red; } }
        public int Green { get { return (int)color.Green; } }
        public int Blue { get { return (int)color.Blue; } }
        public ColorIndex(RGBColor color)
        {
            _index = _globalIndex;
            ++_globalIndex;
            this.color = color;
        }
        public string ToString()
        {
            return string.Format("({0},{1},{2})", Red ,Green,Blue);
        }
    }
    class PolygonInfo
    {
        public Polygon polygon;
        public int index;
        public PolygonInfo(Polygon polygon)
        {
            this.polygon = polygon;
        }
    }
    internal class ColorsToNumber2
    {
        
        #region property
        private const int COUNT_CURVE_POINT = 5;
        private static GeometryFactory _geometryFactory = new GeometryFactory();
        private static Application appRef = new Application();

        private static double MIN_SCALE_VALUE = 1e-6;
        private static double EPSILON = 0.01;

        private static Task update;
        private static string messege = "";
        private static Stack<bool> process = new Stack<bool>();
        private static AutoResetEvent mre = new AutoResetEvent(false);
        private static AutoResetEvent mre2 = new AutoResetEvent(false);
        private static int processCount = 1;
        private static int curTop = 0;
        private static bool terminate = false;

        private static List<ColorIndex> colors = new List<ColorIndex>();

        //private static ConcurrentBag<PolygonInfo> polygons = new ConcurrentBag<PolygonInfo>();
        private static PolygonInfo[] polygons;
        private static BitArray polygonsSet;
        #endregion
        public static void openApp()
        {
            update = Task.Run(() =>
            {
                int p, c;
                Console.CursorVisible = false;
                while (!terminate)
                {

                    Console.SetCursorPosition(0, curTop);
                    Console.WriteLine(messege);
                    Console.Write("[");
                    c = process.Count;
                    p = processCount > 0 ? (int)Math.Round((double)c / processCount * 100) : 100;
                    for (int i = 1; i <= p; i++)
                    {
                        Console.Write('█');
                    }
                    for (int i = p + 1; i <= 100; i++)
                    {
                        Console.Write('░');
                    }
                    Console.Write("]\t" + p + "%");
                    Console.WriteLine(string.Format("\nDone :{0}/{1}", c, processCount));
                    if (c >= processCount)
                    {
                        mre.Set();
                        mre2.WaitOne();
                    }
                }
            });
        }
        private static ColorIndex find(RGBColor color)
        {
            for (int i = 0; i < colors.Count; i++)
                if (CompareColor(color, colors[i].color))
                    return colors[i];
            return null;
        }
        public static bool CompareColor(RGBColor color1, RGBColor color2)
        {
            return color1.Red == color2.Red && color1.Green == color2.Green && color1.Blue == color2.Blue;
        }
        public static int convertRGB(RGBColor color)
        {
            int rgb = (int)color.Red;
            rgb = (rgb << 8) + (int)color.Green;
            rgb = (rgb << 8) + (int)color.Blue;
            return rgb;
        }
        public static Coordinate BeziverFunction(double t, PathPoint P0, PathPoint P3)
        {
            //if (t < 0 || t > 1)
            //    return null;
            double t2 = t * t, t3 = t2 * t;
            double
            x = P0.Anchor[0] * (1 - 3 * t + 3 * t2 - t3) + P0.RightDirection[0] * (3 * t - 6 * t2 + 3 * t3) + P3.LeftDirection[0] * (3 * t2 - 3 * t3) + P3.Anchor[0] * t3,
            y = P0.Anchor[1] * (1 - 3 * t + 3 * t2 - t3) + P0.RightDirection[1] * (3 * t - 6 * t2 + 3 * t3) + P3.LeftDirection[1] * (3 * t2 - 3 * t3) + P3.Anchor[1] * t3;
            return new Coordinate(x, y);
        }
        public static CoordinateList fromAiBeziverCurve(in PathItem pItem)
        {
            double delta = 1 / (double)COUNT_CURVE_POINT;
            int c = pItem.PathPoints.Count + 1;
            CoordinateList result = new CoordinateList();
            for (int i = 1; i < c; i++)
            {
                PathPoint p1 = pItem.PathPoints[i],
                    p3 = pItem.PathPoints[(i + 1) < c ? (i + 1) : 1];

                for (double t = 0; t < 1; t += delta)
                {
                    result.Add(BeziverFunction(t, p1, p3));
                }
            }
            result.Add(result.ElementAt(0).Copy());

            return result;
        }
        public static Polygon from(in CompoundPathItem cpItem)
        {
            int c = cpItem.PathItems.Count;
            List<LinearRing> rings = new List<LinearRing>();
            for (int k = 1; k <= cpItem.PathItems.Count; ++k)
            {
                LinearRing lring = _geometryFactory.CreateLinearRing(fromAiBeziverCurve(cpItem.PathItems[k]).ToArray());
                rings.Add(lring);
            }
            LinearRing outside = rings.MaxBy(ring => ring.Envelope.Area);
            rings.Remove(outside);

            return _geometryFactory.CreatePolygon(outside, rings.ToArray());
        }
        public static Polygon from(in PathItem cpItem)
        {
            return _geometryFactory.CreatePolygon(fromAiBeziverCurve(cpItem).ToArray());
        }
        public static bool reshape(ref Polygon plg, double SCALE)
        {
            if (SCALE < 0)
            {
                return false;
            }
            double midX = (plg.Coordinates[2].X + plg.Coordinates[0].X) / 2,
                   midY = (plg.Coordinates[2].Y + plg.Coordinates[0].Y) / 2,
                   oldWidth2 = SCALE * Math.Abs(plg.Coordinates[2].X - plg.Coordinates[0].X) / 2,
                   oldHeight2 = SCALE * Math.Abs(plg.Coordinates[2].Y - plg.Coordinates[0].Y) / 2;

            Coordinate[] newPlg =
            {
                new Coordinate(midX - oldWidth2,midY + oldHeight2),
                new Coordinate(midX + oldWidth2,midY + oldHeight2),
                new Coordinate(midX + oldWidth2,midY - oldHeight2),
                new Coordinate(midX - oldWidth2,midY - oldHeight2),
                new Coordinate(midX - oldWidth2,midY + oldHeight2),
            };
            plg = _geometryFactory.CreatePolygon(newPlg);

            return true;
        }
        public static bool reshape(ref GroupItem gi, double SCALE)
        {
            if (SCALE < 0)
            {
                gi.Delete();
                return false;
            }
            double midX = gi.Position[0] + gi.Width / 2,
                midY = gi.Position[1] - gi.Height / 2;
            gi.Width = gi.Width * SCALE; gi.Height = gi.Height * SCALE;
            gi.Position = new object[]
            {
                midX - gi.Width / 2,
                midY + gi.Height / 2,
            };

            return true;
        }
        public static void adjust(ref GroupItem gtf, ref Polygon plg, double SCALE)
        {
            Coordinate[] controlBound =
                {
                        new Coordinate(gtf.ControlBounds[0], gtf.ControlBounds[1]),
                        new Coordinate(gtf.ControlBounds[0], gtf.ControlBounds[3]),
                        new Coordinate(gtf.ControlBounds[2], gtf.ControlBounds[3]),
                        new Coordinate(gtf.ControlBounds[2], gtf.ControlBounds[1]),
                        new Coordinate(gtf.ControlBounds[0], gtf.ControlBounds[1])
                    };
            Polygon num = _geometryFactory.CreatePolygon(controlBound);
            double scl = 1, sclw = plg.EnvelopeInternal.Width / num.EnvelopeInternal.Width, sclh = plg.EnvelopeInternal.Height / num.EnvelopeInternal.Height;
            if (sclw < 1 || sclh < 1)
            {
                scl = sclw < sclh ? sclw : sclh;
                reshape(ref num, scl);
            }
            while (true)
            {
                if (plg.Contains(num))
                {
                    break;
                }
                else
                {
                    reshape(ref num, SCALE);
                    scl = scl * SCALE;
                }
                if (scl < MIN_SCALE_VALUE)
                {
                    reshape(ref gtf, -1);
                    return;
                }
            }
            reshape(ref gtf, scl);
        }


        public static void Main(string[] argvs)
        {
            var watch = System.Diagnostics.Stopwatch.StartNew();


            double SCALE = 0.9;
            Application appRef = new Application();
            Document docRef = appRef.ActiveDocument;

            int MinIndex = int.MaxValue, MaxIndex = int.MinValue, k;
            //
            messege = "Loading Compound Path Item";
            processCount = docRef.CompoundPathItems.Count;
            curTop = 0;
            openApp();

            List<CompoundPathItem> cpaths = new List<CompoundPathItem>();
            foreach (CompoundPathItem compoundPathItem in docRef.CompoundPathItems)
            {
                if (compoundPathItem.PathItems[1].Filled && compoundPathItem.Width > EPSILON && compoundPathItem.Height > EPSILON)
                {
                    cpaths.Add(compoundPathItem);
                    k = int.Parse(compoundPathItem.Uuid);
                    if (k < MinIndex) MinIndex = k;
                    if (k > MaxIndex) MaxIndex = k;
                }

                process.Push(true);
            }

            //
            mre.WaitOne();
            messege = "Loading Path Item";
            processCount = docRef.PathItems.Count;
            curTop = curTop + 5;
            process.Clear();
            mre2.Set();

            List<PathItem> paths = new List<PathItem>();
            foreach (PathItem pathItem in docRef.PathItems)
            {
                try
                {
                    if (pathItem.Filled && pathItem.Width > EPSILON && pathItem.Height > EPSILON)
                    {
                        paths.Add(pathItem);
                        k = int.Parse(pathItem.Uuid);
                        if (k < MinIndex) MinIndex = k;
                        if (k > MaxIndex) MaxIndex = k;
                    }
                } catch { }
                process.Push(true);
            }
            k = MaxIndex - MinIndex + 1;
            polygons = new PolygonInfo[k];
            polygonsSet = new BitArray(k, true);

            //
            mre.WaitOne();
            messege = "Loading Polygons";
            processCount = cpaths.Count + paths.Count;
            curTop = curTop + 5;
            process.Clear();
            mre2.Set();

            Parallel.ForEach(cpaths, cpItem =>
            {
                RGBColor rc = cpItem.PathItems[1].FillColor;

                PolygonInfo pinfo = new PolygonInfo(from(cpItem));
                ColorIndex ci = find(rc);
                if (ci is null)
                {
                    ci = new ColorIndex(rc);
                    colors.Add(ci);
                }
                pinfo.index = ci.index;
                polygons[int.Parse(cpItem.Uuid) - MinIndex] = pinfo;

                foreach (PathItem item in cpItem.PathItems)
                {
                    polygonsSet[int.Parse(item.Uuid) - MinIndex] = false;
                }
                process.Push(true);
            });
            Parallel.ForEach(paths, pItem =>
            {
                if (polygonsSet[int.Parse(pItem.Uuid) - MinIndex])
                {

                    RGBColor rc = pItem.FillColor;

                    PolygonInfo pinfo = new PolygonInfo(from(pItem));
                    ColorIndex ci = find(rc);
                    if (ci is null)
                    {
                        ci = new ColorIndex(rc);
                        colors.Add(ci);
                    }

                    pinfo.index = ci.index;
                    polygons[int.Parse(pItem.Uuid) - MinIndex] = pinfo;

                }
                process.Push(true);
            });

            //
            mre.WaitOne();
            messege = "Numbering Polygons";
            processCount = polygons.Length;
            curTop = curTop + 5;
            process.Clear();
            mre2.Set();

            List<GroupItem> colorIndexItems = new List<GroupItem>();
            List<double> colorAreas = new List<double>();
            for (int i = 0; i < colors.Count; ++i)
            {
                colorIndexItems.Add(docRef.GroupItems.Add());
                colorAreas.Add(0.0);
            }


            Parallel.ForEach(polygons, plg =>
            {
                if (plg is object)
                {
                    Point t = plg.polygon.InteriorPoint;

                    TextFrame tf = docRef.TextFrames.Add();
                    tf.Contents = plg.index.ToString();
                    GroupItem gtf = tf.CreateOutline();

                    object[] gtfP =
                    {
                        t.X - gtf.Width/2,
                        t.Y + gtf.Height/2
                    };

                    gtf.Position = gtfP;

                    adjust(ref gtf, ref plg.polygon, SCALE);

                    gtf.Move(colorIndexItems[plg.index -1], AiElementPlacement.aiPlaceInside);


                    colorAreas[plg.index - 1] += plg.polygon.Area;
                    //
                    //
                }
                process.Push(true);
            });

            appRef.ExecuteMenuCommand("Fit Artboard to artwork bounds");

            //double Area = colorAreas.Sum();
            //string textFile = "";
            //for (int i = 0; i < colorAreas.Count; ++i)
            //{
            //    textFile += string.Format("Color {0}: {1} ({2} %)\n", i + 1, colorAreas[i], colorAreas[i] * 100 / Area);
            //}
            //File.WriteAllText("Area.txt", textFile);
            //
            mre.WaitOne();
            messege = "Draw label";
            processCount = colors.Count;
            curTop = curTop + 5;
            process.Clear();
            mre2.Set();

            List<PathItem> squares = new List<PathItem>();
            List<GroupItem> groups = new List<GroupItem>();

            double artX = docRef.Artboards[1].ArtboardRect[0] - 60.0,
                artY = docRef.Artboards[1].ArtboardRect[1];
            for (int i = 0; i < colors.Count; i++)
            {
                PathItem p = docRef.PathItems.Rectangle(0, 0, 30, 30);
                p.FillColor = colors.ElementAt(i).color;
                object[] ps =
                {
                    artX,
                    artY - 60*i
                };
                p.Position = ps;
                squares.Add(p);

                TextFrame tf = docRef.TextFrames.Add();
                tf.Contents = colors.ElementAt(i).index.ToString();
                GroupItem gtf = tf.CreateOutline();

                object[] gtfs =
                {
                    artX + 15 - gtf.Width/2,
                    artY - 15 + gtf.Height/2 - 60*i
                };
                gtf.Position = gtfs;
                groups.Add(gtf);

                process.Push(true);
            }
            appRef.ExecuteMenuCommand("deselectall");
            // export image
            mre.WaitOne();
            messege = "Export Image";
            processCount = colors.Count +1;
            curTop = curTop + 5;
            process.Clear();
            mre2.Set();

            string currentDirectory = Directory.GetCurrentDirectory();
            string storage = currentDirectory + "\\MyFiles\\";

            if (!Directory.Exists(storage))
                Directory.CreateDirectory(storage);
            docRef.SaveAs(storage + docRef.Name.Replace(".ai","_Saved.ai"));
            string storagePNG = currentDirectory + "\\MyFiles\\" + docRef.Name + "_Files\\";
            string imagePage0 = currentDirectory + "\\PDF_cover\\page0.png";

            if (!Directory.Exists(storagePNG))
                Directory.CreateDirectory(storagePNG);
            if (!Directory.Exists(currentDirectory + "\\PDF_cover\\"))
                Directory.CreateDirectory(currentDirectory + "\\PDF_cover\\");


            groups.ForEach(o => o.Delete());
            squares.ForEach(o => o.Delete());

            NoColor noColor = new NoColor();
            ExportForScreensOptionsPNG24 exportForScreensOptions24;
            ExportForScreensItemToExport exportForScreensItemToExport;
            exportForScreensOptions24 = new ExportForScreensOptionsPNG24();
            exportForScreensOptions24.ScaleType = AiExportForScreensScaleType.aiScaleByResolution;
            exportForScreensOptions24.ScaleTypeValue = 300;
            exportForScreensOptions24.Transparency = true;
            exportForScreensItemToExport = new ExportForScreensItemToExport();
            exportForScreensItemToExport.Document = true;
            exportForScreensItemToExport.Artboards = null;

            string saveFolder = docRef.Path.Remove(docRef.Path.LastIndexOf('\\') + 1);
            //for (int i = colors.Count -1; i >= 0; --i)
            //{
            //    squares[i].Selected = true;
            //    appRef.ExecuteMenuCommand("Find Fill Color menu item");
            //    squares[i].Selected = false;
            //    docRef.DefaultFillColor = noColor;
            //    docRef.ExportForScreens(saveFolder + i + ".png", AiExportForScreensType.aiSE_PNG24, exportForScreensOptions24, exportForScreensItemToExport);

            //}
            docRef.ExportForScreens(storagePNG, AiExportForScreensType.aiSE_PNG24, exportForScreensOptions24, exportForScreensItemToExport, colors.Count.ToString("D4"));
            colorIndexItems.ForEach(o => o.Hidden = true);
            process.Push(true);


            docRef.PathItems.Cast<PathItem>().AsParallel().ForAll(o =>
            {
                if (o.Filled && find(o.FillColor) != null)
                {
                    o.Stroked = true;
                    o.StrokeWidth = 0.5;
                    o.StrokeJoin = AiStrokeJoin.aiRoundEndJoin;
                }
            });

            for (int i = colors.Count -1; i >= 0; --i)
            {
                appRef.ExecuteMenuCommand("selectall");
                RGBColor color = colors[i].color;
                docRef.PathItems.Cast<PathItem>().AsParallel().ForAll(o =>
                {
                    if (o.Filled && CompareColor(o.FillColor, color)){
                        o.Selected = false;
                    }
                });
                docRef.DefaultFillColor = noColor;
                colorIndexItems[colors[i].index -1].Hidden = false;
                docRef.ExportForScreens(storagePNG, AiExportForScreensType.aiSE_PNG24, exportForScreensOptions24, exportForScreensItemToExport, i.ToString("D4"));
                appRef.Undo();
                appRef.Undo();
                process.Push(true);
            }
            List<string> imagePaths = Directory.GetFiles(storagePNG, "*.png", SearchOption.AllDirectories).ToList();
            imagePaths.Sort();
            //
            mre.WaitOne();
            messege = "Create PDF";
            processCount = imagePaths.Count +1;
            curTop = curTop + 5;
            process.Clear();
            mre2.Set();
            using (PdfDocument pdf = new PdfDocument())
            {
                PdfPage page0 = pdf.AddPage();
                page0.Size = PdfSharp.PageSize.A4;
                page0.Orientation = PdfSharp.PageOrientation.Portrait;

                // Lấy đối tượng để vẽ
                using (XGraphics gfx = XGraphics.FromPdfPage(page0))
                {

                    // Load ảnh (PNG, JPG, BMP)
                    XImage image = XImage.FromFile(imagePage0);


                    double scaleX = page0.Width / image.PixelWidth;
                    double scaleY = page0.Height / image.PixelHeight;
                    double scale = Math.Min(scaleX, scaleY); // Giữ nguyên tỉ lệ

                    double drawWidth = image.PixelWidth * scale;
                    double drawHeight = image.PixelHeight * scale;

                    // Căn giữa ảnh
                    double x = (page0.Width - drawWidth) / 2;
                    double y = (page0.Height - drawHeight) / 2;

                    // Vẽ ảnh
                    gfx.DrawImage(image, x, y, drawWidth, drawHeight);
                    process.Push(true);
                }

                GlobalFontSettings.UseWindowsFontsUnderWindows = true;
                for (int i = 0; i < imagePaths.Count -1; ++i)
                {
                    string imagePath = imagePaths[i];
                    if (!File.Exists(imagePath))
                    {
                        Console.WriteLine($"File not found: {imagePath}");
                        continue;
                    }

                    // Thêm một trang mới
                    PdfPage page = pdf.AddPage();
                    page.Size = PdfSharp.PageSize.A4;
                    page.Orientation = PdfSharp.PageOrientation.Portrait;


                    // Tạo đối tượng XGraphics để vẽ lên trang
                    using (XGraphics gfx = XGraphics.FromPdfPage(page))
                    {
                        string header = "Color #" + (i + 1).ToString();
                        gfx.DrawString(header, new XFont("Arial", 14), XBrushes.Black,
                                            new XRect(0, 30, page.Width, 30),
                                            new XStringFormat { Alignment = XStringAlignment.Center, LineAlignment = XLineAlignment.Center });
                        // Load hình ảnh
                        using (XImage image = XImage.FromFile(imagePath))
                        {
                            double ratioX = page.Width / image.PixelWidth;
                            double ratioY = page.Height / image.PixelHeight;
                            double scale = Math.Min(ratioX, ratioY);

                            // Kích thước mới
                            double newWidth = image.PixelWidth * scale;
                            double newHeight = image.PixelHeight * scale;

                            // Vị trí căn giữa ảnh
                            double x = (page.Width - newWidth) / 2;
                            double y = (page.Height - newHeight) / 2;

                            // Vẽ ảnh
                            gfx.DrawImage(image, x, y, newWidth, newHeight);
                        }
                    }
                    process.Push(true);
                }
                pdf.Save(saveFolder + docRef.Name + ".pdf");

            }
            using (PdfDocument pdf = new PdfDocument())
            {
                PdfPage page0 = pdf.AddPage();
                page0.Size = PdfSharp.PageSize.A4;
                // Đối tượng để vẽ
                using (XGraphics gfx = XGraphics.FromPdfPage(page0))
                {
                    XFont font = new XFont("Arial", 12);

                    double marginTop = 50;
                    double rowHeight = 40;
                    double rectWidth = 100;
                    double spacing = 10;

                    double totalHeight = marginTop + colors.Count * (rowHeight + spacing);
                    
                    page0.Height = totalHeight;
                    for (int i = 0; i < colors.Count; i++)
                    {
                        ColorIndex color = colors[i];
                        double y = marginTop + i * (rowHeight + spacing);

                        // Chuyển System.Drawing.Color → PdfSharp XColor
                        XColor xcolor = XColor.FromArgb(color.Red, color.Green, color.Blue);
                        XSolidBrush brush = new XSolidBrush(xcolor);

                        // Vẽ hình chữ nhật màu
                        gfx.DrawRectangle(brush, 50, y, rectWidth, rowHeight);

                        // Vẽ text bên cạnh hình chữ nhật
                        double Area = colorAreas.Sum();
                        string info = string.Format("Color {0} = {1}: Area = {2:F4} ({3:F4} %)\n", i + 1, color.ToString(), colorAreas[i], colorAreas[i] * 100 / Area);
                        gfx.DrawString(info, font, XBrushes.Black, 50 + rectWidth + 20, y + rowHeight / 2,
                            new XStringFormat { LineAlignment = XLineAlignment.Center });
                    }
                    process.Push(true);
                }
                pdf.Save(saveFolder + "Colors.pdf");
            }

            mre.WaitOne();
            terminate = true;
            mre2.Set();

            Console.WriteLine("Time Eslapsed :" + watch.Elapsed.TotalSeconds + " seconds");
            watch.Stop();
            Console.WriteLine("Press any key to exit ...");
            Console.ReadKey();
            return;


        }
    }
}
