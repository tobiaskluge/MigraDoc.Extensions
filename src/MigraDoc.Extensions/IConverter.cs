using MigraDoc.DocumentObjectModel;
using System;

namespace MigraDoc.Extensions
{
    public interface IConverter
    {
        Action<DocumentObject> Convert(string contents);
    }
}
