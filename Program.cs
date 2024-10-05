using System.Collections;
using System.Collections.Concurrent;
using System.Xml.XPath;
using Illustrator;
using NetTopologySuite.Geometries;
using NetTopologySuite.Index.HPRtree;

namespace o_w1
{
    class ColorIndex
    {
        private static int _globalIndex = 1;
        private int _index;
        public int index { get { return _index; } }
        public RGBColor color;
        public ColorIndex(RGBColor color)
        {
            _index = _globalIndex;
            ++_globalIndex;
            this.color = color;
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
        private static Stack<bool> process = new Stack<bool> ();
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
                int p,c;
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
                    Console.WriteLine(string.Format("\nDone :{0}/{1}",c, processCount ));
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
            if (SCALE < 0){
                return false;
            }
            double midX = (plg.Coordinates[2].X + plg.Coordinates[0].X) / 2,
                   midY = (plg.Coordinates[2].Y + plg.Coordinates[0].Y) / 2,
                   oldWidth2  = SCALE* Math.Abs(plg.Coordinates[2].X - plg.Coordinates[0].X)/2,
                   oldHeight2 = SCALE * Math.Abs(plg.Coordinates[2].Y - plg.Coordinates[0].Y)/2;

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
            if (SCALE < 0){
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
        public static void adjust(ref GroupItem gtf,ref Polygon plg, double SCALE)
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
            if(sclw < 1 || sclh < 1)
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
                if(scl < MIN_SCALE_VALUE)
                {
                    reshape(ref gtf, -1);
                    return;
                }
            }
            reshape(ref gtf, scl);
        }
        public static void adjust(GroupItem gtf, Polygon plg, Point iterior)
        {
        }
        public static void adjust(GroupItem gtf, Polygon plg, int cWidthGrid, int cHeightGrid)
        {
            Envelope eplg = plg.EnvelopeInternal;
            double deltaX = eplg.Width/ cWidthGrid,
                deltaY = eplg.Height/ cHeightGrid;

            BitArray[] grid = new BitArray[cWidthGrid];
            for (int i = 0; i < cWidthGrid; ++i)
            {
                grid[i] = new BitArray(cHeightGrid, false);
            }
            for (int w = 0; w < cWidthGrid; w++)
            {
                for (int h = 0; h < cHeightGrid; h++)
                {
                    Point p = new Point(eplg.MinX + w*deltaX, eplg.MinY + h*deltaY);
                    if (plg.Contains(p))
                    {
                        grid[w][h] = true;
                    }
                }
            }
            //find Fit rectangle

        }
        public static void findFit(BitArray[] bitArray, int width, int height, List<int> fitX, List<int> fitY)
        {
            int i, ii, j, jj;
            bool check;
            for(i = 0; i < bitArray.Length -width; i++)
            {
                for(j = 0; j < bitArray[i].Length -width; j++)
                {
                    check = true;
                    for(ii = 0; ii < width & check; ii++)
                    {
                        for(jj = 0; jj < height & check; jj++)
                        {
                            if (!bitArray[ii][jj])
                            {
                                check = false;
                            }
                        }
                    }
                    if(check)
                    {
                        fitX.Add(i);
                        fitY.Add(j);
                    }
                }
            }

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
                if (pathItem.Filled && pathItem.Width > EPSILON && pathItem.Height > EPSILON)
                {
                    paths.Add(pathItem);
                    k = int.Parse(pathItem.Uuid);
                    if (k < MinIndex) MinIndex = k;
                    if (k > MaxIndex) MaxIndex = k;
                }
                process.Push(true);
            }
            k = MaxIndex - MinIndex +1;
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

                    adjust(ref gtf,ref plg.polygon, SCALE);
                    //
                    //
                }
                process.Push(true);
            });

            appRef.ExecuteMenuCommand("Fit Artboard to artwork bounds");

            //
            mre.WaitOne();
            messege = "Draw label";
            processCount = colors.Count;
            curTop = curTop + 5;
            process.Clear();
            mre2.Set();

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
                TextFrame tf = docRef.TextFrames.Add();
                tf.Contents = colors.ElementAt(i).index.ToString();
                GroupItem gtf = tf.CreateOutline();

                object[] gtfs =
                {
                    artX + 15 - gtf.Width/2,
                    artY - 15 + gtf.Height/2 - 60*i
                };
                gtf.Position = gtfs;
                process.Push(true);
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
