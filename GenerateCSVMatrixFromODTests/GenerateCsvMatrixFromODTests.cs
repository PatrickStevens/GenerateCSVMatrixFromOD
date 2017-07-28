using Microsoft.VisualStudio.TestTools.UnitTesting;
using GenerateCSVMatrixFromOD;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GenerateCSVMatrixFromOD.Tests
{
  [TestClass]
  public sealed class GenerateCsvMatrixFromODTests
  {
    [TestMethod]
    public void InvalidCommandLineNullArgsTest()
    {
      try
      {
        GenerateCsvMatrixFromOD.Main(null);
        Assert.Fail("Expected exception to be thrown");
      }
      catch (Exception e)
      {
        Assert.AreEqual("Command line format: <layer file name>", e.Message);
      }
    }

    [TestMethod]
    public void InvalidCommandLineNoArgsTest()
    {
      try
      {
        GenerateCsvMatrixFromOD.Main(new string[] {});
        Assert.Fail("Expected exception to be thrown");
      }
      catch (Exception e)
      {
        Assert.AreEqual("Command line format: <layer file name>", e.Message);
      }
    }

    [TestMethod]
    public void InvalidCommandLineTooManyArgsTest()
    {
      try
      {
        GenerateCsvMatrixFromOD.Main(new [] {"arg1", "arg2"});
        Assert.Fail("Expected exception to be thrown");
      }
      catch (Exception e)
      {
        Assert.AreEqual("Command line format: <layer file name>", e.Message);
      }
    }

    [TestMethod]
    public void InvalidCommandLineNoLayerTest()
    {
      const string layerFile = @"c:\temp\Nope.lyr";
      try
      {
        GenerateCsvMatrixFromOD.Main(new[] { layerFile });
        Assert.Fail("Expected exception to be thrown");
      }
      catch (Exception e)
      {
        Assert.AreEqual($"Layer file does not exist: {layerFile}", e.Message);
      }
    }
  }
}