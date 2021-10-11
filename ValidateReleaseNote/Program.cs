using LumenWorks.Framework.IO.Csv;
using Octokit;
using System;
using System.Configuration;
using System.Collections.Generic;
using System.Data.OleDb;
using System.IO;
using System.Text;
using System.Threading.Tasks;

namespace ValidateReleaseNote
{
    class Program
    {
        static void Main(string[] args)
        {
            string executableFilePath = System.Reflection.Assembly.GetExecutingAssembly().Location;
            string directoryPath = Path.GetDirectoryName(executableFilePath);
            string reportFolderPath = directoryPath + "\\" + ConfigurationManager.AppSettings["ValidationFolderName"];
            bool overallSuccess = true;

            if (!Directory.Exists(reportFolderPath))
            {
                Directory.CreateDirectory(reportFolderPath);
            }

            string userName = ConfigurationManager.AppSettings["UserName"];

            StringBuilder errorReport = new StringBuilder("<table cellpadding = '8' cellspacing = '0' style = 'border: 1px solid #ccc;font-size: 9pt;font-family:arial'>");
            errorReport.Append("<tr>");
            errorReport.Append("<th colspan = '4' style='background-color: #FFD700; font-size: 200%;border: 1px solid #ccc'>Release Note Validation - Executed By : " + userName + "</th>");
            errorReport.Append("</tr>");
            errorReport.Append("<tr>");
            errorReport.Append("<th style='background-color: #B8DBFD;border: 1px solid #ccc'>FileName</th>");
            errorReport.Append("<th style='background-color: #B8DBFD;border: 1px solid #ccc'>Developer</th>");
            errorReport.Append("<th style='background-color: #B8DBFD;border: 1px solid #ccc'>Status</th>");
            errorReport.Append("<th style='background-color: #B8DBFD;border: 1px solid #ccc'>Message</th>");
            errorReport.Append("</tr>");

            try
            {
                Console.WriteLine(" Release note artifact names and GIT check in reference validation is in progress... ");

                string password = ConfigurationManager.AppSettings["Password"];
                string token = ConfigurationManager.AppSettings["token"];
                string gitUrl = ConfigurationManager.AppSettings["GitUrl"];
                string repositoryOwner = ConfigurationManager.AppSettings["RepositoryOwner"];
                string repository = ConfigurationManager.AppSettings["Repository"];

                if (string.IsNullOrEmpty(password) && string.IsNullOrEmpty(token))
                    throw new Exception("Both password and token can not be blank at the same time.");

                string filePath = directoryPath + "\\" + ConfigurationManager.AppSettings["ReleaseNoteName"];

                var csvTable = new System.Data.DataTable();
                var excelTable = new System.Data.DataTable();
                List<ReleaseNoteFormat> releaseNote = new List<ReleaseNoteFormat>();
                if (Path.GetExtension(filePath).ToLower() == ".csv")
                {
                    using (var csvReader = new CsvReader(new StreamReader(File.OpenRead(filePath)), true))
                    {
                        csvTable.Load(csvReader);
                    }

                    for (int i = 0; i < csvTable.Rows.Count; i++)
                    {
                        if (!string.IsNullOrEmpty(csvTable.Rows[i][0].ToString()) || !string.IsNullOrEmpty(csvTable.Rows[i][1].ToString()) || !string.IsNullOrEmpty(csvTable.Rows[i][3].ToString()))
                        {
                            releaseNote.Add(new ReleaseNoteFormat { LineNumber = csvTable.Rows[i][0].ToString(), CROrDefectReference = csvTable.Rows[i][1].ToString(), FileName = csvTable.Rows[i][3].ToString(), ChangeType = csvTable.Rows[i][4].ToString(), GitCheckInReference = csvTable.Rows[i][8].ToString(), Developer = csvTable.Rows[i][2].ToString() });
                        }
                    }
                }
                else if (Path.GetExtension(filePath).ToLower() == ".xlsx" || Path.GetExtension(filePath).ToLower() == ".xls")
                {
                    excelTable = ReadExcel(filePath, Path.GetExtension(filePath).ToLower());

                    for (int i = 1; i < excelTable.Rows.Count; i++)
                    {
                        if (!string.IsNullOrEmpty(excelTable.Rows[i][0].ToString()) || !string.IsNullOrEmpty(excelTable.Rows[i][1].ToString()) || !string.IsNullOrEmpty(excelTable.Rows[i][3].ToString()))
                        {
                            releaseNote.Add(new ReleaseNoteFormat { LineNumber = excelTable.Rows[i][0].ToString(), CROrDefectReference = excelTable.Rows[i][1].ToString(), FileName = excelTable.Rows[i][3].ToString(), ChangeType = excelTable.Rows[i][4].ToString(), GitCheckInReference = excelTable.Rows[i][8].ToString(), Developer = excelTable.Rows[i][2].ToString() });
                        }
                    }

                }

                GitHub github = new GitHub();
                GitHubClient gitHubClient = github.GetGitHubClient(userName, password, token, gitUrl);

                foreach (ReleaseNoteFormat noteFormat in releaseNote)
                {
                    try
                    {
                        if (string.IsNullOrEmpty(noteFormat.ChangeType))
                            throw new Exception("Change Type can not be blank.");

                        if (!string.IsNullOrEmpty(noteFormat.FileName) && !string.IsNullOrEmpty(noteFormat.GitCheckInReference) && !noteFormat.GitCheckInReference.ToLower().Contains("n/a") && !noteFormat.GitCheckInReference.ToLower().Contains("na"))
                        {
                            string[] references = noteFormat.GitCheckInReference.Split('/');
                            string gitCheckInReference = references[references.Length - 1];
                            Task<List<GitHubCommitFile>> files = github.GetAllFiles(gitHubClient, repositoryOwner, repository, gitCheckInReference);
                            files.Wait(5000);
                            bool invalid = true;
                            foreach (GitHubCommitFile file in files.Result)
                            {
                                string[] names = file.Filename.Split('/');
                                string fileName = names[names.Length - 1];
                                if (noteFormat.FileName.ToLower().Equals(fileName.ToLower()))
                                {
                                    invalid = false;
                                    Console.WriteLine(noteFormat.FileName + " exists in GITHUB checkin reference: " + noteFormat.GitCheckInReference);
                                    errorReport.Append("<tr>");
                                    errorReport.Append("<td style='width:120px;border: 1px solid #ccc'>" + noteFormat.FileName + "</td>");
                                    errorReport.Append("<td style='width:120px;border: 1px solid #ccc'>" + noteFormat.Developer + "</td>");
                                    errorReport.Append("<td style='width:120px;border: 1px solid #ccc'>Success</td>");
                                    errorReport.Append("<td> </td>");
                                    errorReport.Append("</tr>");

                                    break;
                                }

                            }

                            if (invalid)
                            {
                                overallSuccess = false;
                                Console.WriteLine(noteFormat.FileName + " does not exist in GITHUB checkin reference: " + noteFormat.GitCheckInReference);
                                errorReport.Append("<tr>");
                                errorReport.Append("<td style='width:120px;border: 1px solid #ccc'>" + noteFormat.FileName + "</td>");
                                errorReport.Append("<td style='width:120px;border: 1px solid #ccc'>" + noteFormat.Developer + "</td>");
                                errorReport.Append("<td bgcolor='red' style='width:120px;border: 1px solid #ccc'>Failed</td>");
                                errorReport.Append("<td style='width:120px;border: 1px solid #ccc'>For " + noteFormat.CROrDefectReference + ", artifact does not exist under the GITHUB checkin reference.</td>");
                                errorReport.Append("</tr>");
                            }
                        }
                        else
                        {
                            if (string.IsNullOrEmpty(noteFormat.FileName))
                            {
                                overallSuccess = false;
                                Console.WriteLine("FileName does not exist in the release note for line number: " + noteFormat.LineNumber);
                                errorReport.Append("<tr>");
                                errorReport.Append("<td style='width:120px;border: 1px solid #ccc'>" + noteFormat.FileName + "</td>");
                                errorReport.Append("<td style='width:120px;border: 1px solid #ccc'>" + noteFormat.Developer + "</td>");
                                errorReport.Append("<td bgcolor='red' style='width:120px;border: 1px solid #ccc'>Failed</td>");
                                errorReport.Append("<td style='width:120px;border: 1px solid #ccc'>For " + noteFormat.CROrDefectReference + ", FileName does not exist under the release note.</td>");
                                errorReport.Append("</tr>");
                            }

                            if (string.IsNullOrEmpty(noteFormat.GitCheckInReference) && noteFormat.ChangeType.ToLower().Equals("code"))
                            {
                                overallSuccess = false;
                                Console.WriteLine("Git check in reference does not exist in the release note for line number: " + noteFormat.LineNumber);
                                errorReport.Append("<tr>");
                                errorReport.Append("<td style='width:120px;border: 1px solid #ccc'>" + noteFormat.FileName + "</td>");
                                errorReport.Append("<td style='width:120px;border: 1px solid #ccc'>" + noteFormat.Developer + "</td>");
                                errorReport.Append("<td bgcolor='red' style='width:120px;border: 1px solid #ccc'>Failed</td>");
                                errorReport.Append("<td style='width:120px;border: 1px solid #ccc'>For " + noteFormat.CROrDefectReference + ", Git Check In Reference does not exist under the release note.</td>");
                                errorReport.Append("</tr>");
                            }
                        }
                    }
                    catch (Exception exception)
                    {
                        overallSuccess = false;
                        string exceptionMessage = !string.IsNullOrEmpty(exception.InnerException.Message.ToString()) ? exception.InnerException.Message.ToString() : exception.Message.ToString();
                        Console.WriteLine("Exception occurred for line number : " + noteFormat.LineNumber + ", " + exceptionMessage);
                        errorReport.Append("<tr>");
                        errorReport.Append("<td style='width:120px;border: 1px solid #ccc'>" + noteFormat.FileName + "</td>");
                        errorReport.Append("<td style='width:120px;border: 1px solid #ccc'>" + noteFormat.Developer + "</td>");
                        errorReport.Append("<td bgcolor='red' style='width:120px;border: 1px solid #ccc'>Failed</td>");
                        errorReport.Append("<td style='width:120px;border: 1px solid #ccc'>For " + noteFormat.CROrDefectReference + ", " + exceptionMessage + " </td>");
                        errorReport.Append("</tr>");
                    }
                }

                if (overallSuccess)
                {
                    Console.WriteLine("\nValidation Status: Succeeded");
                    errorReport.Append("<tr><td colspan = '3' style='background-color: #008000; font-size: 150%;border: 1px solid #ccc'>Overall Validation: Succeeded</td></tr>");
                }
                else
                {
                    Console.WriteLine("\nValidation Status: Failed");
                    errorReport.Append("<tr><td colspan = '3' style='background-color: #FF0000; font-size: 150%;border: 1px solid #ccc'>Overall Validation: Failed</td></tr>");
                }

                errorReport.Append("</table>");

                File.WriteAllText(reportFolderPath + "\\" + "ValidationStatusReport-" + DateTime.Now.ToString("dd-MM-yyyy-HH-mm-ss-ffff") + ".htm", errorReport.ToString());
                Console.WriteLine("\nExecution is completed. Press any key to close the console...");
                Console.ReadKey();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception occurred: " + ex.Message.ToString());
                Console.WriteLine("\nValidation Status: Failed");
                errorReport.Append("<tr>");
                errorReport.Append("<td colspan = '3' style='width:120px;border: 1px solid #ccc'>Error: " + ex.Message.ToString() + "</td>");
                errorReport.Append("</tr>");
                errorReport.Append("<tr><td colspan = '3' style='background-color: #FF0000; font-size: 150%;border: 1px solid #ccc'>Overall Validation: Failed</td></tr>");
                errorReport.Append("</table>");
                File.WriteAllText(reportFolderPath + "\\" + "ValidationStatusReport-" + DateTime.Now.ToString("dd-MM-yyyy-HH-mm-ss-ffff") + ".htm", errorReport.ToString());
                Console.WriteLine("\nExecution is completed. Press any key to close the console...");
                Console.ReadKey();
            }


        }

        public static System.Data.DataTable ReadExcel(string fileName, string fileExt)
        {
            string conn = string.Empty;
            string sheetName = ConfigurationManager.AppSettings["SheetName"];
            System.Data.DataTable dtexcel = new System.Data.DataTable();

            if (fileExt.CompareTo(".xls") == 0)
                conn = @"provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileName + ";Extended Properties='Excel 8.0;HRD=Yes;IMEX=1';"; //for below excel 2007  
            else
                conn = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties='Excel 12.0;HDR=NO';"; //for above excel 2007  
            using (OleDbConnection con = new OleDbConnection(conn))
            {
                try
                {
                    OleDbDataAdapter oleAdpt = new OleDbDataAdapter("select * from [" + sheetName + "$]", con);
                    oleAdpt.Fill(dtexcel);
                }
                catch(Exception ex) { }
            }
            return dtexcel;
        }
    }
}
