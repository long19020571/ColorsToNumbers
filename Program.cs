using System;
using System.Collections;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Illustrator;
using NetTopologySuite.Geometries;
using NetTopologySuite.Index.HPRtree;
using NetTopologySuite.Shape;

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
        private const int COUNT_CURVE_POINT = 20;
        private static GeometryFactory _geometryFactory = new GeometryFactory();
        private static Application appRef = new Application();

        private static ConcurrentBag<ColorIndex> colors = new ConcurrentBag<ColorIndex>();

        private static ConcurrentBag<PolygonInfo> polygons = new ConcurrentBag<PolygonInfo>();
        #endregion
        public void openApp() { }
        private static ColorIndex find(RGBColor color)
        {
            foreach (ColorIndex c in colors)
            {
                if (CompareColor(color, c.color)) return c;
            }
            return null;
        }
        public static bool CompareColor(RGBColor color1, RGBColor color2)
        {
            return color1.Red == color2.Red & color1.Green == color2.Green & color1.Blue == color2.Blue;
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

                for (double t = 0; t <= 1; t += delta)
                {
                    result.Add(BeziverFunction(t, p1, p3));
                }
            }
            result.Add(result.ElementAt(0).Copy());

            return result;
        }
        public static Polygon from(in CompoundPathItem cpItem)
        {
            Console.WriteLine("\tStart : from");
            int c = cpItem.PathItems.Count;
            List<LinearRing> rings = new List<LinearRing>();
            for (int k = 1; k <= cpItem.PathItems.Count; ++k)
            {
                LinearRing lring = _geometryFactory.CreateLinearRing(fromAiBeziverCurve(cpItem.PathItems[k]).ToArray());
                rings.Add(lring);
            }
            LinearRing outside = rings.MaxBy(ring => ring.Envelope.Area);
            rings.Remove(outside);

            Console.WriteLine("\tEnd : from");
            return _geometryFactory.CreatePolygon(outside, rings.ToArray());
        }
        public static Polygon from(in PathItem cpItem)
        {
            return _geometryFactory.CreatePolygon(fromAiBeziverCurve(cpItem).ToArray());
        }
        public static Coordinate BeziverFunction(double t, PathPoint P0, PathPoint P3)
        {
            if (t < 0 || t > 1)
                return null;
            double t2 = t * t, t3 = t2 * t;
            double
            x = P0.Anchor[0] * (1 - 3 * t + 3 * t2 - t3) + P0.RightDirection[0] * (3 * t - 6 * t2 + 3 * t3) + P3.LeftDirection[0] * (3 * t2 - 3 * t3) + P3.Anchor[0] * t3,
            y = P0.Anchor[1] * (1 - 3 * t + 3 * t2 - t3) + P0.RightDirection[1] * (3 * t - 6 * t2 + 3 * t3) + P3.LeftDirection[1] * (3 * t2 - 3 * t3) + P3.Anchor[1] * t3;
            return new Coordinate(x, y);
        }
        public static void resizeItem(GroupItem gItem, double scale)
        {
            double m_x = gItem.Position[0] + gItem.Width / 2, m_y = gItem.Position[1] - gItem.Height / 2;

            gItem.Height = gItem.Height * scale;
            gItem.Width = gItem.Width * scale;

            object[] p =
            {
                m_x - gItem.Width/2,
                m_y - gItem.Height/2
            };
            gItem.Position = p;
        }
        public static void adjust(GroupItem gtf, Polygon plg, double SCALE, Point iterior)
        {
            while (true)
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
                if (plg.Contains(num))
                {
                    break;
                }
                else
                {
                    gtf.Width = gtf.Width * SCALE;
                    gtf.Height = gtf.Height * SCALE;

                    object[] tfP2 = {
                            iterior.X - gtf.Width / 2,
                            iterior.Y + gtf.Height / 2
                        };

                    gtf.Position = tfP2;
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
                cpaths.Add(compoundPathItem);
            List<PathItem> paths = new List<PathItem>();
            foreach (PathItem pathItem in docRef.PathItems)
                paths.Add(pathItem);
            ConcurrentBag<PathItem> excludePathItems = new ConcurrentBag<PathItem>();

            Parallel.ForEach(cpaths, cpItem =>
            {
                Console.WriteLine("CompoundPathItem:" + cpItem.Name + "-" + cpItem.Uuid);
                RGBColor rc = null;
                bool catchE = false;
                try
                {
                    rc = cpItem.PathItems[1].FillColor;
                }
                catch
                {
                    catchE = true;
                }
                if (!catchE)
                {
                    PolygonInfo pinfo = new PolygonInfo(from(cpItem));
                    ColorIndex ci = find(rc);
                    if (ci == null)
                    {
                        ci = new ColorIndex(rc);
                        colors.Add(ci);
                    }
                    pinfo.index = ci.index;
                    polygons.Add(pinfo);

                    foreach (PathItem item in cpItem.PathItems)
                    {
                        excludePathItems.Add(item);
                        Console.WriteLine("~PathItem:" + item.Name + "-" + item.Uuid);
                    }
                } else
                {
                    Console.WriteLine("\tThrow Excepion in " + cpItem.Name + "-" + cpItem.Uuid);
                }
            });
            Parallel.ForEach(paths, pItem =>
            {
                if (!excludePathItems.Contains(pItem))
                {
                    Console.WriteLine("PathItem:" + pItem.Name + "-" + pItem.Uuid);

                    RGBColor rc = null;
                    bool catchE = false;
                    try
                    {
                        rc = pItem.FillColor;
                    }
                    catch
                    {
                        catchE = true;
                    }
                    if (!catchE)
                    {
                        Console.WriteLine("\t- Adding ...");
                        PolygonInfo pinfo = new PolygonInfo(from(pItem));
                        ColorIndex ci = find(rc);
                        if (ci == null)
                        {
                            ci = new ColorIndex(rc);
                            colors.Add(ci);
                        }
                        pinfo.index = ci.index;
                        polygons.Add(pinfo);
                    } else
                    {

                        Console.WriteLine("\tThrow Excepion in " + pItem.Name + "-" + pItem.Uuid);
                    }
                }
            });

            paths = null;
            cpaths = null;
            excludePathItems = null;
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

                adjust(gtf, plg.polygon, 0.9, t);
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

            return;
            //gtf.Position = tfP;


            //int count = docRef.PathItems.Count;
            //List<PathItem> pList = new List<PathItem>();
            //foreach (PathItem item in docRef.PathItems)
            //{
            //    pList.Add(item);
            //}

            //for (int i = 0; i < pList.Count; i++)
            //{

            //    PathItem p = pList[i];
            //    Polygon plg = gF.CreatePolygon(fromAiBeziverCurve(p).ToArray());
            //    Point iterior = plg.InteriorPoint;

            //    Console.WriteLine("Setting :" + i + " at: " + iterior.X + ", "+ iterior.Y);

            //    TextFrame tf = docRef.TextFrames.Add();
            //    tf.Contents = i.ToString();
            //    GroupItem gtf = tf.CreateOutline();

            //    object[] tfP =
            //    {
            //        iterior.X - gtf.Width/2,
            //        iterior.Y + gtf.Height/2
            //    };

            //    gtf.Position = tfP;

            //    while (true)
            //    {
            //        Console.WriteLine("\tPositining :" + i + " Width:" + gtf.Width);
            //        Coordinate[] controlBound =
            //        {
            //            new Coordinate(gtf.ControlBounds[0], gtf.ControlBounds[1]),
            //            new Coordinate(gtf.ControlBounds[0], gtf.ControlBounds[3]),
            //            new Coordinate(gtf.ControlBounds[2], gtf.ControlBounds[3]),
            //            new Coordinate(gtf.ControlBounds[2], gtf.ControlBounds[1]),
            //            new Coordinate(gtf.ControlBounds[0], gtf.ControlBounds[1])
            //        };
            //        Polygon num = gF.CreatePolygon(controlBound);
            //        if (plg.Contains(num))
            //        {
            //            break;
            //        }
            //        else
            //        {
            //            gtf.Width = gtf.Width * SCALE;
            //            gtf.Height = gtf.Height * SCALE;

            //            object[] tfP2 = {
            //                iterior.X - gtf.Width / 2,
            //                iterior.Y + gtf.Height / 2
            //            };

            //            gtf.Position = tfP2;
            //        }
            //    }
            //}



        }
    }
}
