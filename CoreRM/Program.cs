using System;
using System.Linq;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Reflection;
using System.Security;
using Microsoft.Extensions.CommandLineUtils;

namespace CoreRM
{
    class Program
    {
        static void Main(string[] args)
        {
            var app = new CommandLineApplication
            {
                Name = "CoreRM - WinRM via .NET Core",
                Description = "WinRM with .NET Core",
                ExtendedHelpText = "Example Usage: dotnet CoreRM.dll -u user_name"
            };

            app.HelpOption("-?|-h|--help");

            app.VersionOption("-v|--version", () =>
            {
                return string.Format("Version {0}", Assembly.GetEntryAssembly().GetCustomAttribute<AssemblyInformationalVersionAttribute>().InformationalVersion);
            });

            var userNameOption = app.Option("-u|--user <user_name>",
                    "Username",
                    CommandOptionType.SingleValue);

            var targetOption = app.Option("-t|--target <target>",
                    "Target Hostname or IP Address",
                    CommandOptionType.SingleValue);

            var portOption = app.Option("-p|--port <port>",
                    "Port (Default: 5985)",
                    CommandOptionType.SingleValue);

            var pathOption = app.Option("--path <path>",
                    "Path (Default: /wsman)",
                    CommandOptionType.SingleValue);

            app.OnExecute(() =>
            {
                if (userNameOption.HasValue() && targetOption.HasValue())
                {
                    var user = userNameOption.Value();
                    var target = targetOption.Value();
                    var password = new SecureString();

                    Console.Write("Password for {0}: ", user);
                    password = GetPassword();
                    Console.WriteLine("");

                    string shell = "http://schemas.microsoft.com/powershell/Microsoft.PowerShell";
                    var targetWsMan = new Uri(string.Format("http://{0}:{1}{2}", target, portOption.HasValue() ? portOption.Value() : "5985", pathOption.HasValue() ? pathOption.Value() : "/wsman"));
                    var cred = new PSCredential(user, password);
                    var connectionInfo = new WSManConnectionInfo(targetWsMan, shell, cred)

                    {
                        OperationTimeout = 4 * 60 * 1000, // 4 minutes.
                        OpenTimeout = 1 * 60 * 1000,
#if DEBUG
                        AuthenticationMechanism = System.Management.Automation.Runspaces.AuthenticationMechanism.Basic
#endif
                    };
                    using (var runSpace = RunspaceFactory.CreateRunspace(connectionInfo))
                    {
                        var cmd = "";
                        runSpace.Open();
                        do
                        {
                            var p = runSpace.CreatePipeline();
                            if (!string.IsNullOrEmpty(cmd))
                            {
                                p.Commands.AddScript(cmd);
                                try
                                {
                                    var results = p.Invoke();

                                    foreach (var o in results)
                                    {
                                        Console.WriteLine(o.ToString());
                                    }
                                }
                                catch (Exception ex)
                                {
                                    var prevColor = Console.ForegroundColor;
                                    Console.ForegroundColor = ConsoleColor.Red;
                                    Console.WriteLine(ex.Message);
                                    Console.ForegroundColor = prevColor;
                                }
                            }
                            p = runSpace.CreatePipeline();
                            p.Commands.AddScript("$(pwd).Path");
                            var dir = p.Invoke().Select(x => x.ToString()).FirstOrDefault() ?? "";

                            Console.Write("PS {0}@{1}:{2}>", user, target, dir);
                            cmd = Console.ReadLine();
                        } while (cmd != "exit");

                        runSpace.Disconnect();
                    }
                }
                else
                {
                    app.ShowHint();
                }
                return 0;
            });

            try
            {
                app.Execute(args);
            }
            catch (CommandParsingException ex)
            {
                var prevColor = Console.ForegroundColor;
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(ex.Message);
                Console.ForegroundColor = prevColor;
            }
            catch (Exception ex)
            {
                var prevColor = Console.ForegroundColor;
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(ex.Message);
                Console.ForegroundColor = prevColor;
            }
        }

        private static SecureString GetPassword()
        {
            var password = new SecureString();
            ConsoleKeyInfo keyInfo;
            do
            {
                keyInfo = Console.ReadKey(true);
                if (keyInfo.Key != ConsoleKey.Backspace && keyInfo.Key != ConsoleKey.Enter)
                {
                    password.AppendChar(keyInfo.KeyChar);
                    //Console.Write("*");
                }
                else
                {
                    if (keyInfo.Key == ConsoleKey.Backspace && password.Length > 0)
                    {
                        password.RemoveAt(password.Length - 1);
                        //Console.Write("\b \b");
                    }
                }
            }
            while (keyInfo.Key != ConsoleKey.Enter);

            return password;
        }
    }
}
