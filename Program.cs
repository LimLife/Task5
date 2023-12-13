using TaskToXLSX10._12._23.Application;
using TaskToXLSX10._12._23.Data;

IFileEditor xmlFileReader = new FileEditor();
var app = new Application()
{
    _editor = xmlFileReader
};

app.Input();


