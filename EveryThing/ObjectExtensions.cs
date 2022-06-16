using System;
using System.Collections.Generic;
using System.Linq;
using Tekla.Structures.Geometry3d;

namespace EveryThing
{
    public static class ObjectExtensions
    {
        public static void AddIfNotContains(this List<Point> pointList, Point point)
        {
            if (!pointList.Contains(point, new EqualityComparerPoint()))
            {
                pointList.Add(point);
            }
        }
        public static void AddIfNotContains(this List<LineSegment> segmentList, LineSegment segment)
        {
            if (!segmentList.Contains(segment, new EqualityComparerLineSegment()))
            {
                segmentList.Add(segment);
            }
        }
    }
}
