using Cartes;
using MiTools;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Windows.Forms;

namespace ImagesOCR
{
    static class Program
    {
        static private CartesObj cartes = new CartesObj();

        [STAThread]
        static void Main()
        {
            DateTime start;
            RPAParameters parameters = new RPAParameters(), output = null;
            RPAWin32Component notepad = null, notepadclose = null, notepadeditor = null,
                              notepadnosave = null, notepaddialog = null;
            String workingFile, language = "eng", imageCloseButton;

            start = DateTime.Now;
            workingFile = Environment.CurrentDirectory;
            imageCloseButton = workingFile + "\\closebutton.bmp";
            cartes.open(workingFile + "\\library\\notepad2.rpa");
            try
            {
                notepad = cartes.GetComponent<RPAWin32Component>("$Notepad");
                notepadclose = cartes.GetComponent<RPAWin32Component>("$NotepadClose");
                notepadeditor = cartes.GetComponent<RPAWin32Component>("$NotepadEditor");
                notepadnosave = cartes.GetComponent<RPAWin32Component>("$NotepadNoSave");
                notepaddialog = cartes.GetComponent<RPAWin32Component>("$NotepadDialog");
                if (notepadeditor.componentexist(0) == 0)
                {
                    cartes.run("notepad.exe");
                    notepadeditor.waitforcomponent(30);
                }
                notepad.ReSize(890, 600);
                notepad.Move(1, 5);
                notepad.focus();
                notepad.SaveRectPartToFile(notepadclose.x - notepad.x,
                                           notepadclose.y - notepad.y,
                                           notepadclose.width, notepadclose.height,
                                           imageCloseButton);
                parameters.item[0] = imageCloseButton;
                output = notepad.FindPicture(parameters);
                if (output.item[0] == "1")
                {
                    notepadeditor.TypeFromClipboard("RESULT : " + output.item[0] + "\r\n" +
                                                    "X      : " + output.item[1] + "\r\n" +
                                                    "Y      : " + output.item[2] + "\r\n" +
                                                    "WIDTH  : " + output.item[3] + "\r\n" +
                                                    "HEIGHT : " + output.item[4] + "\r\n" +
                                                    "INDEX  : " + output.item[5] + "\r\n");
                    // I use the OCR directly on the screen with zoom
                    MessageBox.Show(notepadeditor.RecognitionRatio(language, 1.3, 1));
                    // I use the OCR in an image file
                    Thread.Sleep(2000);
                    notepadeditor.focus();
                    notepadeditor.SaveRectToFile(imageCloseButton);
                    cartes.Execute("$OCR = new OCR;\r\n" +
                                   "ShowMessage($OCR.run(\"" + imageCloseButton + "\", \"" + language + "\"));\r\n");
                    // Closing...
                    notepad.clickon(int.Parse(output.item[1]) + int.Parse(output.item[3]) / 2,
                                    int.Parse(output.item[2]) + int.Parse(output.item[4]) / 2, 1); // I use the "ClickOn" function directly with the coordinates
                    Thread.Sleep(1000);
                    cartes.reset(notepadnosave.api());
                    notepadnosave.waitforcomponent(10);
                    notepaddialog.SaveRectPartToFile(notepadnosave.x - notepaddialog.x,
                                                     notepadnosave.y - notepaddialog.y,
                                                     notepadnosave.width, notepadnosave.height,
                                                     imageCloseButton);
                    notepaddialog.ClickOnImage(true, imageCloseButton);
                }
                cartes.RegisterIteration(start, "ok", "<task>Put your trace here in xml for your swarm</task>", 1);
                cartes.forensic("This is a trace for the swarm log, and the Windows event viewer.");
                cartes.balloon("This example has opened the notepad and has shown how to use image recognition and OCR.");
            }
            finally
            {
                cartes.close();
                MessageBox.Show("End");
            }
        }
    }
}
