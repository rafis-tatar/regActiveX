using System;

namespace regActiveX
{
    internal class Program
    {
        private static int Main(string[] args)
        {
            var check = false;
            var isRegistr = false;
            var logo = true;
            var isOleControl = true;
            string filePath = null;
            
            if (args == null || args.Length == 0)
            {
                ShowCopirite();
                ShowHelp();
                return -1;
            }

            foreach (var argument in args)
            {
                if (argument.Length < 2)
                {
                    ShowHelp();
                    return -1;
                }

                if (argument.StartsWith("/"))
                {
                    switch (argument.Substring(1, 1))
                    {
                        case "c":
                            check = true;
                            break;
                        case "r":
                            isRegistr = true;
                            break;
                        case "u":
                            isRegistr = false;
                            break;
                        case "n":
                            logo = false;
                            break;
                        case "x":
                            isOleControl = false;
                            break;
                        default:
                            ShowHelp();
                            return -1;
                    }
                }
                else
                {
                    filePath = argument;
                }
            }

            if (logo)
            {
                ShowCopirite();
            }
            
            if (string.IsNullOrWhiteSpace(filePath))
            {
                ShowHelp();
                return -1;
            }


            var conRegistration = new ComRegistration();
            try
            {
                if (check)
                {
                    return conRegistration.CheckClassCOM(filePath) ? 1 : 0;
                }
                if (isRegistr)
                    conRegistration.RegisterCOM(filePath, isOleControl);
                else
                    conRegistration.UnregisterCOM(filePath);

            }
            catch (Exception exception)
            {
                Console.WriteLine($"ERROR - {exception.Message}");
                return 0;
            }
            return 1;
        }

        private static void ShowCopirite()
        {
            Console.WriteLine();
            Console.WriteLine("Программа регистрации типов/классов (OleControls) библиотеки .NET Framework для КИС АРМИТС");
            Console.WriteLine("(C) Группа АРМИТС. Все права защищены");
            Console.WriteLine();
        }
        private static void ShowHelp()
        {
            Console.WriteLine("Синтаксис: regActiveX Наименование файла библиотеки [Параметры]");
            Console.WriteLine("Параметры:");
            Console.WriteLine("\t/r Регистрация dll");
            Console.WriteLine("\t/u Отмена регистрации dll");
            Console.WriteLine("\t/c Проверка регистрация dll");
            Console.WriteLine("\t/n Не выводить логотип");
            Console.WriteLine("\t/x Регистрация как NOT OleControl (зарегистрированный класс)");
            Console.WriteLine("\t/h Это сообщение");
        }
    }
}