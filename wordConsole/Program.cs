using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;


namespace wordConsole
{
    class Program
    {


        static void Main(string[] args)
        {
            Methods methods = new Methods();
            Console.WriteLine(methods.openWord());
            bool cicle = true;
            String res;
            while (cicle)
            {
                Console.WriteLine("S=search, C=count, Q=quit");
                res = Console.ReadLine();
                switch (res)
                {
                    case "q":
                        Console.WriteLine(methods.quitWord());
                        cicle = false;
                        break;
                    case "s":
                        res = Console.ReadLine();
                        Console.WriteLine(methods.findNextInDoc(res));
                        break;
                    case "c":
                        res = Console.ReadLine();
                        Console.WriteLine(methods.countTextTimes(res));
                        break;
                }
            }

            Console.ReadLine();
        }
    }
}

