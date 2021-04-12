using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;
using System.Reflection;

namespace WindowsFormsApplication1
{
    class groupDoc
    {
        private Object _missingObj = System.Reflection.Missing.Value;
        private Object _trueObj = true;
        private Object _falseObj = false;

        private Word._Application _application;
        private Word._Document _document;

        public groupDoc(string templatePath, bool startVisible)
        {
            //создаем обьект приложения word
            _application = new Word.Application();

            // создаем путь к файлу используя имя файла
            Object templatePathObj = "C:\\Documents and Settings\\iruchi\\Application Data\\Microsoft\\Шаблоны\\Normal.dotm";
        }
    }
}
