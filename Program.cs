using System.Collections;
using System.Collections.Concurrent;
using Illustrator;
using NetTopologySuite.Geometries;

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

        private static ConcurrentBag<ColorIndex> colors = new ConcurrentBag<ColorIndex>();

        private static ConcurrentBag<PolygonInfo> polygons = new ConcurrentBag<PolygonInfo>();
        #endregion
        public void openApp() { }
        private static ColorIndex find(RGBColor color)
        {
            for (int i = 0; i < colors.Count; i++)
                if (CompareColor(color, colors.ElementAt(i).color))
                    return colors.ElementAt(i);
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
            //if (!plg.IsRectangle)
            //    return false;
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
            //if (gi == null)
            //    return false;
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
        public static void adjust(GroupItem gtf, Polygon plg, double SCALE)
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
            double scl = 1;
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

            List<CompoundPathItem> cpaths = new List<CompoundPathItem>();
            foreach (CompoundPathItem compoundPathItem in docRef.CompoundPathItems)
            {
                if (compoundPathItem.PathItems[1].Filled)
                    cpaths.Add(compoundPathItem);
            }
            List<PathItem> paths = new List<PathItem>();
            foreach (PathItem pathItem in docRef.PathItems)
            {
                if (pathItem.Filled)
                    paths.Add(pathItem);
            }
            ConcurrentDictionary<int,PathItem> excludePathItems = new ConcurrentDictionary<int, PathItem>();

            Parallel.ForEach(cpaths, cpItem =>
            {
                Console.WriteLine("CompoundPathItem:" + cpItem.Name + "-" + cpItem.Uuid);
                RGBColor rc = cpItem.PathItems[1].FillColor;
                
                PolygonInfo pinfo = new PolygonInfo(from(cpItem));
                ColorIndex ci = find(rc);
                if (ci is null)
                {
                    ci = new ColorIndex(rc);
                    colors.Add(ci);
                }
                pinfo.index = ci.index;
                polygons.Add(pinfo);

                foreach (PathItem item in cpItem.PathItems)
                {
                    excludePathItems.TryAdd(int.Parse(item.Uuid), item);
                    Console.WriteLine("~PathItem:" + item.Name + "-" + item.Uuid);
                }
            });
            Parallel.ForEach(paths, pItem =>
            {
                if (!excludePathItems.ContainsKey(int.Parse(pItem.Uuid)))
                {
                    Console.WriteLine("PathItem:" + pItem.Name + "-" + pItem.Uuid);

                    RGBColor rc = pItem.FillColor;

                    PolygonInfo pinfo = new PolygonInfo(from(pItem));
                    ColorIndex ci = find(rc);
                    if (ci is null)
                    {
                        ci = new ColorIndex(rc);
                        colors.Add(ci);
                    }

                    pinfo.index = ci.index;
                    polygons.Add(pinfo);

                }
            });

            Console.WriteLine(watch.Elapsed.TotalSeconds);
            Console.WriteLine("Numbering polygon");

            Parallel.ForEach(polygons, plg =>
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

                adjust(gtf, plg.polygon, SCALE);
                //
                //

            });
            
            appRef.ExecuteMenuCommand("Fit Artboard to artwork bounds");

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
            }

            watch.Stop();
            Console.WriteLine(watch.Elapsed.TotalSeconds);
            Console.ReadKey();
            return;


        }
    }
}
