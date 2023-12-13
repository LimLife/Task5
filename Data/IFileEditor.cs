namespace TaskToXLSX10._12._23.Data
{
    public interface IFileEditor : IXmlFileReader, IXMLFileWriterCustomer
    {
        public void SetPathToFileString(string path);
    }
}
