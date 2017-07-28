using System;

namespace GenerateCSVMatrixFromOD
{
  internal sealed class ArcGISLicensingException : Exception
  {
    public ArcGISLicensingException(string message) : base(message)
    {
    }
  }
}