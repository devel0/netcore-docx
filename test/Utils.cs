using Xunit;
using System.Linq;
using System;
using static System.Math;
using DocumentFormat.OpenXml.Wordprocessing;

namespace SearchAThing.DocX.Tests
{

    public static partial class Ext
    {

        public static void AssertEqualsTol(this double actual, double tol, double expected, string userMessage = "")
        {
            if (!expected.EqualsTol(tol, actual))
                throw new Xunit.Sdk.AssertActualExpectedException(expected, actual, userMessage);
        }

        
    }

}

