namespace OPCBrowser
{
    using System;
    using System.Runtime.InteropServices;
    using OPCAutomation;

    class Program
    {        
        static void Main(string[] args)
        {
            switch (args.Length) {
                case 0:
                    ShowOpcServerNames("localhost");
                    break;
                case 1:
                    ShowOpcServerNames(args[0]);
                    break;
                default:
                    ShowOpcTagsList(args[0], args[1]);
                    break;
            }
        }

        /// <summary>
        /// Показывает все доступные OPC-сервера на хосте.
        /// </summary>
        /// <param name="aHost">Хост.</param>
        private static void ShowOpcServerNames(string aHost)
        {            
            try {                
                var serversNames = new OPCServer().GetOPCServers(aHost) as Array;
                if (serversNames != null && serversNames.Length > 0) {
                    Console.WriteLine("Список всех OPC-серверов по адресу {0}:", aHost);
                    for (var i = 1; i <= serversNames.Length; ++i) {
                        Console.WriteLine("{0} {1}", i, serversNames.GetValue(i));
                    }
                }
                else {
                    Console.WriteLine("OPC-серверов по адресу {0} не найдено.", aHost);
                }
            }
            catch (COMException ex) {
                Console.WriteLine("COM-ошибка при просмотре списка OPC-серверов: " + ex.Message);                
            } 
        }

        /// <summary>
        /// Подключается к OPC-серверу и запускает рекурсивный обход его содержимого.
        /// </summary>
        /// <param name="aHost">Хост.</param>
        /// <param name="aOpcServerName">Имя OPC-сервера.</param>
        private static void ShowOpcTagsList(string aHost, string aOpcServerName)
        {            
            var opcServer = new OPCServer();
            try {
                opcServer.Connect(aOpcServerName, aHost);
            }
            catch (COMException ex) {
                Console.WriteLine("Не удалось подключиться к OPC-серверу по адресу {0}. Ошибка: {1}.",
                    aHost, ex.Message);
                return;
            }

            ShowOpcTagsRecurcive(opcServer.CreateBrowser());
        }

        /// <summary>
        /// Рекурсивно обходит содержимое OPC-сервера.
        /// </summary>
        /// <param name="aBrowser">Браузер OPC-сервера.</param>
        /// <param name="aPrefix">Префикс для построения дерева.</param>
        private static void ShowOpcTagsRecurcive(OPCBrowser aBrowser, string aPrefix = "")
        {
            var isError = false;
            try {
                aBrowser.ShowBranches();
            }
            catch (Exception ex) {
                Console.WriteLine("Ошибка при отображении ветки.");
                isError = true;
            }

            if (!isError) {
                foreach (var item in aBrowser) {
                    Console.WriteLine(aPrefix + "Ветвь: " + item);
                    aBrowser.MoveDown(item.ToString());
                    ShowOpcTagsRecurcive(aBrowser, aPrefix + "-");
                    aBrowser.MoveUp();
                }
            }

            isError = false;
            try {
                aBrowser.ShowLeafs();
            }
            catch (Exception ex) {
                Console.WriteLine("Ошибка при отображении листа.");
                isError = true;
            }

            if (!isError) {
                foreach (var item in aBrowser) {
                    Console.WriteLine(aPrefix + "Лист:" + item);
                }
            }
        }
    }
}
