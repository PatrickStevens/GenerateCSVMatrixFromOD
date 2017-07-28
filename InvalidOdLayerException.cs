using System;

namespace GenerateCSVMatrixFromOD
{
  internal sealed class InvalidOdLayerException : Exception
  {
    public InvalidOdLayerException(string message) : base(message)
    {
    }
  }
}