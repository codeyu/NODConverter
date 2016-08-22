namespace NODConverter
{
    public interface IDocumentFormatRegistry
    {

        DocumentFormat GetFormatByFileExtension(string extension);

        DocumentFormat GetFormatByMimeType(string extension);

    }
}
