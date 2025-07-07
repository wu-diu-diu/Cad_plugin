using System;
using System.Collections.Generic;
using System.Linq;
using MIConvexHull;
using System.Windows;

namespace BoundingRectangle
{
    // Fix: Implement IVertex2D and add a parameterless constructor
    public class Vertex : IVertex2D
    {
        public double[] Position { get; set; }

        public double X => Position[0];
        public double Y => Position[1];

        public Vertex()
        {
            Position = new double[2];
        }

        public Vertex(double x, double y)
        {
            Position = new double[] { x, y };
        }
    }

    public class Program
    {
        public static List<Point> CalculateBoundingRectangle(List<double[]> rawPoints)
        {
            List<Vertex> points2D = rawPoints.Select(p => new Vertex(p[0], p[1])).ToList();
            var hull = ConvexHull.Create2D(points2D);

            // Fix: Access the Result property directly, as Points is not defined in IList<Vertex>.  
            var convexPoints = hull.Result.Select(p => new Point(p.Position[0], p.Position[1])).ToList();

            var rectangle = GetMinimumBoundingBox(convexPoints);

            return rectangle;
        }

        static List<Point> GetMinimumBoundingBox(List<Point> convexHull)
        {
            double minArea = double.MaxValue;
            List<Point> bestBox = null;

            for (int i = 0; i < convexHull.Count; i++)
            {
                var p1 = convexHull[i];
                var p2 = convexHull[(i + 1) % convexHull.Count];
                var edge = new Vector(p2.X - p1.X, p2.Y - p1.Y);
                var angle = Math.Atan2(edge.Y, edge.X);

                var rotated = convexHull.Select(p =>
                {
                    double x = p.X * Math.Cos(-angle) - p.Y * Math.Sin(-angle);
                    double y = p.X * Math.Sin(-angle) + p.Y * Math.Cos(-angle);
                    return new Point(x, y);
                }).ToList();

                double minX = rotated.Min(p => p.X), maxX = rotated.Max(p => p.X);
                double minY = rotated.Min(p => p.Y), maxY = rotated.Max(p => p.Y);
                double area = (maxX - minX) * (maxY - minY);

                if (area < minArea)
                {
                    minArea = area;
                    bestBox = new List<Point>
                                {
                                    new Point(minX, minY),
                                    new Point(maxX, minY),
                                    new Point(maxX, maxY),
                                    new Point(minX, maxY)
                                }.Select(p =>
                                {
                                    double x = p.X * Math.Cos(angle) - p.Y * Math.Sin(angle);
                                    double y = p.X * Math.Sin(angle) + p.Y * Math.Cos(angle);
                                    return new Point(x, y);
                                }).ToList();
                }
            }
            return bestBox;
        }
    }
}
