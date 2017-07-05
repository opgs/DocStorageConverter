using Microsoft.Office.Word.Server.Conversions;
using Microsoft.SharePoint;
using System;
using System.Collections.ObjectModel;
using System.Threading;

namespace OPG.DocStorageConverter
{
    class Converter
    {
        static void Main(string[] args)
        {
            Console.WriteLine("OPG DocStorageConverter");
            Console.WriteLine("");

            Console.WriteLine("Connecting to:\t\t" + Settings.siteUrl);
            using (SPSite spSite = new SPSite(Settings.siteUrl))
            {
                SPFolder folderToConvert = spSite.RootWeb.GetFolder(Settings.fileconvert);
                ConversionJob job = new ConversionJob(Settings.wordAutomationServiceName);
                job.UserToken = spSite.UserToken;
                job.Settings.UpdateFields = true;
                job.Settings.OutputFormat = SaveFormat.PDF;
                job.Settings.FixedFormatSettings.UsePDFA = true;

                foreach(SPFile file in folderToConvert.Files)
                {
                    if(Files.HasExtension(file.Name, new string[]{ ".docx", ".docm", ".dotx", ".dotm", ".doc", ".dot", ".rtf", ".mht", ".mhtml", ".xml" }))
                    {
                        string filePath = (Settings.siteUrl + "/" + file.Url);
                        Console.WriteLine("Found:\t\t\t" + filePath);
                        using (SPWeb web = spSite.OpenWeb())
                        {
                            if(web.GetFile(filePath).Exists)
                            {
                                Console.WriteLine("Already Exists:\t\t" + Files.StripExtension(filePath) + ".pdf");
                                Console.WriteLine("Deleting:\t\t" + Files.StripExtension(filePath) + ".pdf");
                                SPFile eFile = web.GetFile(Files.StripExtension(filePath) + ".pdf");
                                eFile.Delete();
                                eFile.Update();
                            }
                        }
                        job.AddFile(filePath, Files.StripExtension(filePath) + ".pdf");
                    }                    
                }

                try
                {
                    job.Start();
                }
                catch (InvalidOperationException)
                {
                    Console.WriteLine("Done:\t\t\tNo files to convert");
                    return;
                }

                Console.WriteLine("\t\t\tConversion job started");
                ConversionJobStatus status = new ConversionJobStatus(Settings.wordAutomationServiceName, job.JobId, null);
                Console.WriteLine("Job length:\t\t" + status.Count);

                while (true)
                {
                    Thread.Sleep(5000);
                    status = new ConversionJobStatus(Settings.wordAutomationServiceName, job.JobId, null);
                    if (status.Count == status.Succeeded + status.Failed)
                    {
                        Console.WriteLine("Completed:\t\tSuccessful: " + status.Succeeded + ", Failed: " + status.Failed);
                        ReadOnlyCollection<ConversionItemInfo> failedItems = status.GetItems(ItemTypes.Failed);
                        foreach (ConversionItemInfo failedItem in failedItems)
                        {
                            Console.WriteLine("Failed converting:\t" + failedItem.InputFile);
                            Console.WriteLine(failedItem.ErrorMessage);
                        }
                        Console.WriteLine("\t\t\tSetting meta on files that successfully converted");
                        ReadOnlyCollection<ConversionItemInfo> convertedItems = status.GetItems(ItemTypes.Succeeded);
                        SPSecurity.RunWithElevatedPrivileges(delegate ()
                        {
                            using (SPWeb web = spSite.OpenWeb())
                            {
                                web.AllowUnsafeUpdates = true;
                                foreach (ConversionItemInfo convertedItem in convertedItems)
                                {
                                    SPFile inFile = web.GetFile(convertedItem.InputFile);
                                    SPFile outFile = web.GetFile(convertedItem.OutputFile);
                                    try
                                    {
                                        SPListItem inListItem = inFile.Item;
                                        SPListItem outListItem = outFile.Item;
                                        Console.WriteLine("Set metadata on:\t" + outFile.Url);
                                        foreach (SPField field in inListItem.Fields)
                                        {
                                            try
                                            {
                                                if (outListItem.Fields.ContainsField(field.InternalName) == true && field.ReadOnlyField == false && field.InternalName != "Attachments" && field.InternalName != "Name")
                                                {
                                                    outListItem[field.InternalName] = inListItem[field.InternalName];
                                                    Console.WriteLine("Setting field:\t\t" + field.InternalName);

                                                }
                                            }
                                            catch (Exception e)
                                            {
                                                Console.WriteLine("Failed to set field:\t" + field.InternalName + " : " + e.Message);
                                            }
                                        }
                                        outListItem.Update();
                                        outFile.Update();
                                    }
                                    catch (Exception e)
                                    {
                                        Console.WriteLine("Failed to set on:\t" + outFile.Url + " from : " + inFile.Url);
                                        Console.WriteLine(e.Message);
                                    }
                                }
                                web.AllowUnsafeUpdates = false;
                            }
                        });
                        Console.WriteLine("\t\t\tDeleting only items that successfully converted");
                        foreach (ConversionItemInfo convertedItem in convertedItems)
                        {
                            Console.WriteLine("Deleting item:\t\tName:" + convertedItem.InputFile);
                            folderToConvert.Files.Delete(convertedItem.InputFile);
                        }
                        break;
                    }
                    Console.WriteLine("In progress:\t\tSuccessful: " + status.Succeeded + ", Failed: " + status.Failed);
                }

                Console.WriteLine("Done:\t\t\tFinished");
            }
        }
    }
}
