/*
 * Windows update force installer
 * A Simple tool used to fix errors such as 0x80070003 that are acoused by corrupted Windows updates
 * How to use: Download the missing/corrupted update file from https://www.catalog.update.microsoft.com/Home.aspx
 * 
 * 
 * 
 * 
 * Created by: Arttu Mahlakaarto (https://github.com/amahlaka) on 29.11.2019
 */
using System;
using System.Security.Principal;
using Mono.Options;
using System.Diagnostics;
namespace ConsoleApp1
{
    public class UpdateFile
    {
        public string Filename { get; set; }
        public string Path { get; set; }
        public string[] Files { get; set; }
        public string Cabinet { get; set; }
        public bool Debug { get; set; }
        public UpdateFile()
        {
        }
    }
    class Program
    {
        static void Main(string[] args)
        {

            bool is_elevated;
            using (WindowsIdentity identity = WindowsIdentity.GetCurrent())
            {
                WindowsPrincipal principal = new WindowsPrincipal(identity);
                is_elevated = principal.IsInRole(WindowsBuiltInRole.Administrator);
            }
            if (!is_elevated)
            {
                Console.WriteLine("This application requires Adminstrator privileges");
                Console.WriteLine("Error: Access Denied (5)");
                System.Environment.Exit(5);
            }

            bool help = false;
            bool repair = false;
            UpdateFile updateFile = new UpdateFile();
            OptionSet option_set = new OptionSet()
                .Add("?|help|h", "Prints out the options.", option => help = option != null)
                .Add("file=|f=", "REQUIRED: .msu file to install", option => updateFile.Filename = option)
                .Add("path=|p=", "REQUIRED: Temp directory", option => updateFile.Path = option)
                .Add("debug", "Debug, disable dism execution, print the commands instead", option => updateFile.Debug = option != null)
                .Add("repair", "Optional: Perform a DISM RestoreHealth after installing the update", option => repair = option != null);
            option_set.Parse(args);
            if (updateFile.Path == null || updateFile.Filename == null || help)
            {
                PrintHelp(option_set);
                System.Environment.Exit(0);
            }
            if (!System.IO.File.Exists(updateFile.Filename))
            {
                Console.WriteLine("That file does not exist, aborting");
                System.Environment.Exit(15);
            }
            if (System.IO.Directory.Exists(updateFile.Path))
            {
                Console.WriteLine("That Folder already exists!");
                string[] tempFiles = System.IO.Directory.GetFiles(updateFile.Path);
                if (tempFiles.Length != 0)
                {
                    Console.WriteLine("The folder " + updateFile.Path + " contains the following files:");
                    foreach (var item in tempFiles)
                    {
                        Console.WriteLine(item);
                    }
                    Console.Write("Do you want to delete all the files in this folder? Y/N:");
                    string UserInput = Console.ReadLine();
                    switch (UserInput)
                    {
                        case "Y":
                        case "y":
                            Console.Clear();
                            Console.WriteLine("Deleting...");
                            DeleteFiles(tempFiles);
                            break;
                        case "N":
                        case "n":
                            Console.WriteLine("Aborting...");
                            System.Environment.Exit(1);
                            break;
                        default:
                            Console.WriteLine("Invalid option, aborting just to be safe!");
                            System.Environment.Exit(1);
                            break;
                    }
                }



            }
            else
            {
                Console.WriteLine("Target folder does not exists, Creating it now...");
                System.IO.Directory.CreateDirectory(updateFile.Path);
            }
            ExpandFile(updateFile);
            ApplyUpdate(updateFile);
            if (repair)
            {
                RestoreHealth(updateFile);
            }
            Console.WriteLine("All actions completed!");
            Console.WriteLine("Press any key to exit");
            Console.ReadKey();

        }
        static void DeleteFiles(string[] files)
        {
            // Simple (Possibly unneccesary) Wrapper to delete files in the folder.
            foreach (var item in files)
            {
                if (System.IO.File.Exists(item))
                {
                    Console.Write("Deleting file: " + item + "...");
                    System.IO.File.Delete(item);
                    Console.WriteLine(" - Done!");
                }
            }
        }
        static void ExpandFile(UpdateFile file)
        {
            // Extract the .cab file from the Standalone update installer. (DISM cannot patch a .msu file to a currently running Windows image)

            Process expand = new Process();
            expand.StartInfo.FileName = "expand.exe";
            expand.StartInfo.Arguments = "/f:* \"" + file.Filename + "\" \"" + file.Path + "\"";
            expand.StartInfo.UseShellExecute = false;
            expand.StartInfo.RedirectStandardOutput = true;
            expand.Start();
            expand.WaitForExit();
            expand.Dispose();
            
            file.Files = System.IO.Directory.GetFiles(file.Path);
            if (file.Files.Length < 1)
            {
                Console.WriteLine("There seems to have been a error expanding the files");
            }
            Console.WriteLine("Expansion completed!");
            foreach (var item in file.Files)
            {
                if (System.IO.Path.GetExtension(item) == ".cab" && item.Contains("KB"))
                {
                    Console.WriteLine("Cabinet file is: " + item);
                    file.Cabinet = item;
                }
            }

        }
        static void ApplyUpdate(UpdateFile file)
        {
            // Apply the update directly to currently running Windows Image.
            if(file.Debug == true)
            {
                Console.WriteLine("DEBUG: dism.exe /Online /Add-Package /PackagePath:\"" + file.Cabinet + "\"");
                return;
            }
            Console.Clear();
            Process dism = new Process();
            dism.StartInfo.FileName = "dism.exe";
            dism.StartInfo.Arguments = "/Online /Add-Package /PackagePath:\"" + file.Cabinet + "\"";
            dism.Start();
            dism.WaitForExit();
            dism.Dispose();

        }
        static void RestoreHealth(UpdateFile file)
        {
            if (file.Debug == true)
            {
                Console.WriteLine("DEBUG: dism.exe /Online /Cleanup-Image /RestoreHealth");
                return;
            }
            Process dism = new Process();
            dism.StartInfo.FileName = "dism.exe";
            dism.StartInfo.Arguments = "/Online /Cleanup-Image /RestoreHealth";
            dism.Start();
            dism.WaitForExit();
            dism.Dispose();
        }
        static void PrintHelp(OptionSet options)
        {
            Console.WriteLine("Usage: command [OPTIONS]");
            Console.WriteLine("Download the Correct .msu installer file from: https://www.catalog.update.microsoft.com/Home.aspx");
            options.WriteOptionDescriptions(Console.Out);

        }
    }
}
