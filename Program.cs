using PuppeteerSharp;
using System;
using System.ComponentModel.DataAnnotations;
using System.Globalization;
using System.IO;
using System.Threading.Tasks;
using static System.Net.Mime.MediaTypeNames;

class Program
{
    static async Task Main(string[] args)
    {
        short dia = x;
        short mes = y;
        short ano = z;

        var browserFetcher = new BrowserFetcher();
        await browserFetcher.DownloadAsync();

        var launchOptions = new LaunchOptions
        {
            Headless = false
        };

        using (var browser = await Puppeteer.LaunchAsync(launchOptions))
        using (var page = await browser.NewPageAsync())
        {
            await page.SetViewportAsync(new ViewPortOptions
            {
                Width = 1920,
                Height = 1080,
                DeviceScaleFactor = 1
            });

            //Dados do login e da senha
            string txt = @"C:\Users\lelopes\OneDrive - TRIBUNAL DE JUSTICA DE PERNAMBUCO\Documents\TXT\file.txt";

            Account conta = Account.LoadData(txt);

            string email = conta.Email;
            string password = conta.Password;

            // Login and password
            await page.GoToAsync("https://login.microsoftonline.com/");

            await page.WaitForSelectorAsync("input[type='email']");
            await page.TypeAsync("input[type='email']", email);

            await page.ClickAsync("#idSIButton9");
            await Task.Delay(2000);

            await page.WaitForSelectorAsync("input[type='password']");
            await page.TypeAsync("input[type='password']", password);

            await page.ClickAsync("#idSIButton9");
            await Task.Delay(1000);

            await page.WaitForSelectorAsync("#idBtn_Back");
            await page.ClickAsync("#idBtn_Back");
            await Task.Delay(2000);

            // OneDrive
            /*await page.GoToAsync("https://tjpejus-my.sharepoint.com/");
            await page.WaitForSelectorAsync("a[title='Meus arquivos']");
            await page.ClickAsync("a[title='Meus arquivos']");
            await Task.Delay(2000);

            await page.WaitForSelectorAsync("button[title='Chamados']");
            await page.ClickAsync("button[title='Chamados']");
            await Task.Delay(2000);

            // Navigate to specific date
            await page.WaitForSelectorAsync($"button[title='{ano}']");
            await page.ClickAsync($"button[title='{ano}']");
            await Task.Delay(2000);

            await page.WaitForSelectorAsync($"button[title='{mes}']");
            await page.ClickAsync($"button[title='{mes}']");
            await Task.Delay(2000);

            await page.WaitForSelectorAsync($"button[title='{dia}']");
            await page.ClickAsync($"button[title='{dia}']");
            await Task.Delay(2000);*/

            for (int j = dia; j <= 31; j++)
            {
                await page.GoToAsync($"https://straight_to_link.com.br/url...");
                
                await Task.Delay(5000);

                string folder;
                if (j < 10)
                {
                    folder = @"D:\Comunix\Gravações\" + ano + "-" + mes + "-0" + j;
                }
                else
                {
                    folder = @"D:\Comunix\Gravações\" + ano + "-" + mes + "-" + j;
                }
                Console.WriteLine(folder);

                // Get list of audio files in folder
                string[] files = Directory.GetFiles(folder);
                Console.WriteLine($"Quantidade total de chamados no dia {j}/{mes}/{ano}: {files.Length} áudios.\n\n");

                string folderChamados = $"C://Users//file_path//" + ano + "//" + mes + "//" + j;
                string[] filesChamados = Directory.GetFiles(folderChamados);

                for (int i = filesChamados.Length; i < files.Length; i++)
                {
                    // Click "Adicionar novo" button
                    await page.WaitForSelectorAsync("button[title='Adicionar novo']", new WaitForSelectorOptions { Timeout = 0 });
                    await page.ClickAsync("button[title='Adicionar novo']");
                    await Task.Delay(1500);

                    // Click "Documento do Word" button
                    await page.WaitForSelectorAsync("button[title='Documento do Word']");
                    var link = await page.QuerySelectorAsync("button[title='Documento do Word']");
                    await Task.Delay(1500);

                    var newPagePromise = new TaskCompletionSource<Page>();
                    EventHandler<TargetChangedArgs> targetChangedHandler = null;

                    // Define a handler for the target changed event
                    targetChangedHandler = async (sender, e) =>
                    {
                        var target = e.Target;
                        if (target.Type == TargetType.Page)
                        {
                            var page2 = await target.PageAsync();

                            newPagePromise.SetResult((Page)page2);
                            browser.TargetChanged -= targetChangedHandler;
                        }
                    };

                    // Subscribe to the target changed event
                    browser.TargetChanged += targetChangedHandler;

                    // Click on the link
                    await link.ClickAsync();
                    var page2 = await newPagePromise.Task;
                    await Task.Delay(5000);

                    DateTime date = DateTime.Now;
                    Console.WriteLine("\t\t\tHora atual: " + date.Hour + ":" + date.Minute + ":" + date.Second);

                    var audio = Path.GetFileName(files[i]);
                    var doc = Path.GetFileNameWithoutExtension(audio);
                    var ID = doc.Substring(0, 12) + '.' + doc.Substring(12);

                    // Access the iframe within the new page
                    var iframeHandle = await page2.WaitForSelectorAsync("iframe", new WaitForSelectorOptions { Timeout = 0 });
                    var frame = await iframeHandle.ContentFrameAsync();
                    await Task.Delay(3000);

                    // Write ID to the Word document
                    await frame.WaitForSelectorAsync("#PagesContainer > div");
                    await frame.TypeAsync("#PagesContainer > div", ID + " - ");
                    await Task.Delay(5000);

                    // Open the document title field
                    await frame.WaitForSelectorAsync("#documentTitle");
                    await frame.ClickAsync("#documentTitle");
                    await Task.Delay(5000);

                    // Change the document title
                    await frame.WaitForSelectorAsync("#CommitNewDocumentTitle");
                    await frame.TypeAsync("#CommitNewDocumentTitle", ID);
                    await Task.Delay(5000);

                    // Click on the overflow menu to open options
                    await frame.WaitForSelectorAsync("#DictationTranscriptionSplit > button:nth-child(2) > span", new WaitForSelectorOptions { Timeout = 0 });
                    await frame.ClickAsync("#DictationTranscriptionSplit > button:nth-child(2) > span");
                    await Task.Delay(5000);

                    // Click on the transcription button
                    await frame.WaitForSelectorAsync("button[name='Transcrever']");
                    await frame.ClickAsync("button[name='Transcrever']");
                    await Task.Delay(5000);

                    // Access the iframe within the new page
                    var iframeHandle2 = await frame.WaitForSelectorAsync("iframe", new WaitForSelectorOptions { Timeout = 0 });
                    var frame2 = await iframeHandle2.ContentFrameAsync();

                    // Upload the audio file
                    Console.Write($"{i + 1}ª Transcrição:\nCaminho do arquivo: {files[i]}\nID: {ID}\n");
                    await Task.Delay(5000);

                    await frame2.WaitForSelectorAsync("input[type=file]");

                    var fileInput = await frame2.QuerySelectorAsync("input[type=file]");
                    await fileInput.UploadFileAsync(files[i]);

                    // Click the "addToDocument" button
                    await frame2.WaitForSelectorAsync("button[id='addToDocument']", new WaitForSelectorOptions { Timeout = 0 });
                    await Task.Delay(2000);
                    await frame2.ClickAsync("button[id='addToDocument']");

                    await Task.Delay(2000);

                    await frame2.WaitForSelectorAsync("button[name='Com palestrantes e carimbo de data/hora']");
                    await frame2.ClickAsync("button[name='Com palestrantes e carimbo de data/hora']");

                    await Task.Delay(5000);

                    Console.WriteLine("\nTranscrição finalizada!\n\n");

                    await page2.CloseAsync();

                    await Task.Delay(2000);
                }

                Console.WriteLine("\nTodas as transcrições do dia " + j + " foram finalizadas!\n");
            }
            await browser.CloseAsync();
        }
    }
}
