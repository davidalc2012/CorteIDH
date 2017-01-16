using System;
using Word = Microsoft.Office.Interop.Word;

namespace wordConsole
{
    class Methods
    {

        Word.Application application;
        Word.Document doc;

        public String openWord()
        {
            try
            {
                application = new Microsoft.Office.Interop.Word.Application();
                doc = new Microsoft.Office.Interop.Word.Document();
            }
            catch(Exception e)
            {
                return e.Message;
            }

            object missing = System.Reflection.Missing.Value;
            object readOnly = false;
            object isVisible = true;
            object fileName = @"C:\Users\David\Desktop\WpfApplication1\WpfApplication1\Título.docx";

            //object fileName = "http://www.corteidh.or.cr/docs/casos/articulos/seriec_330_esp.docx";

            try
            {
                application.Visible = true;
                doc = application.Documents.Open(ref fileName, ref missing, ref readOnly, ref missing, ref missing, ref missing, 
                                                ref missing, ref missing, ref missing, ref missing, ref missing, ref isVisible, 
                                                ref missing, ref missing, ref missing, ref missing);
                doc.Activate();
            }
            catch (Exception e)
            {
                return e.Message;
            }
            return "WORD DOCUMENT OPEN";
        }

        public bool findNextInDoc(String toFind)
        {
            try
            {
                Object text = toFind;
                application.Selection.Find.ClearFormatting();
                object missing = System.Reflection.Missing.Value;

                if (application.Selection.Find.Execute(ref text,
                    ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing))
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch (Exception e)
            {
                System.Console.WriteLine(e.Message);
                return false;
            }

        }

        public int countTextTimes(String toFind)
        {
            object what = Word.WdGoToItem.wdGoToLine;
            object wich = Word.WdGoToDirection.wdGoToAbsolute;
            object missing = 1;
            application.Selection.GoTo(what, wich, missing).Select();
            
            int count=0;
            while (findNextInDoc(toFind))
            {
                count++;
            }
            return count;
        }

        public String quitWord()
        {
            Object saveChanges = Word.WdSaveOptions.wdDoNotSaveChanges;
            Object originalFormat = Word.WdOriginalFormat.wdOriginalDocumentFormat;
            Object routeDocument = false;
            String error="";
            try
            {
                doc.Close(saveChanges, originalFormat, routeDocument);
            }
            catch (Exception e)
            {
                error= e.Message;
            }
            try
            {
                application.Quit(saveChanges, originalFormat, routeDocument);
            }
            catch (Exception e)
            {
                return error+" "+e.Message;
            }
            if (error!= ""){
                return error;
            }
            return "Word closed";
        }

        public void tests()
        {
            object start= doc
            Word.Range rng= doc.
            application.Selection.SetRange(, 15);
        }
        
    }
}

/*object what = Word.WdGoToItem.wdGoToLine;
            object wich = Word.WdGoToDirection.wdGoToAbsolute;
            object missing = 15;
            application.Selection.GoTo(what, wich, missing).Select();
            //System.Console.WriteLine(application.Selection.Range.Information[Word.WdInformation.wdFirstCharacterLineNumber]);

            object unit = Word.WdUnits.wdLine;
            object extend = Word.WdMovementType.wdExtend;
            application.Selection.MoveDown(unit, 2, extend);

            //application.Selection.SetRange(0, 15);
            Word.Characters wchar = application.Selection.Characters;
            System.Console.WriteLine(wchar.ToString());
            System.Console.WriteLine(application.Selection.Text);
            System.Console.ReadLine();

    */