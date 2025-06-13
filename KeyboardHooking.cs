using System.Runtime.InteropServices;
using System.Windows.Forms;
using System;
using Microsoft.Office.Interop.Word;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Word = Microsoft.Office.Interop.Word;

namespace PasteEquation
{
    class WordFunctions
    {
        // Count existing inline equations in the current document
        private static int CountExistingEquations(Word.Document doc)
        {
            int n = 0;
            try
            {
                foreach (Word.InlineShape iSh in doc.InlineShapes)
                {
                    if (iSh.OLEFormat != null)
                    {
                        string progId = iSh.OLEFormat.ProgID;
                        if (progId == "Equation.3" || progId == "Word.Equation.8") n++;
                    }
                }
            }
            catch
            {
                // If we can't count equations, assume moderate load
                return 50;
            }
            return n;
        }        private static void Paste(string input)
        {
            try
            {
                Clipboard.SetText(input);
                Range currentRange = Globals.Main.Application.Selection.Range;
                currentRange.Paste();
                Clipboard.Clear();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Paste error: {ex.Message}");
                // Attempt to clear clipboard on error to prevent corruption
                try { Clipboard.Clear(); } catch { }
                throw; // Re-throw so higher level can handle
            }
        }

        private static bool SplitAndPaste(string input)
        {
            try
            {
                // For extremely large content, split into smaller chunks first
                if (input.Length > 50000)
                {
                    System.Diagnostics.Debug.WriteLine($"Large content detected ({input.Length} chars), using pre-chunking");
                    return ProcessLargeContent(input);
                }

                return ProcessNormalContent(input);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"SplitAndPaste error: {ex.Message}");
                return false;
            }
        }

        // Handle extremely large content by splitting it first
        private static bool ProcessLargeContent(string input)
        {
            try
            {
                // Split large content into manageable chunks (preserve sentence structure)
                var chunks = SplitIntoChunks(input, 25000);
                
                System.Diagnostics.Debug.WriteLine($"Split large content into {chunks.Count} chunks");
                
                bool success = true;
                for (int i = 0; i < chunks.Count; i++)
                {
                    System.Diagnostics.Debug.WriteLine($"Processing chunk {i + 1}/{chunks.Count}");
                    
                    if (!ProcessNormalContent(chunks[i]))
                    {
                        success = false;
                    }
                    
                    // Extra pause between chunks for very large content
                    if (i < chunks.Count - 1)
                    {
                        System.Threading.Thread.Sleep(500);
                        System.Windows.Forms.Application.DoEvents();
                        
                        // Force garbage collection every few chunks
                        if (i % 2 == 1)
                        {
                            GC.Collect();
                            GC.WaitForPendingFinalizers();
                        }
                    }
                }
                
                return success;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ProcessLargeContent error: {ex.Message}");
                return false;
            }
        }

        // Split content into chunks while preserving sentence boundaries
        private static List<string> SplitIntoChunks(string input, int maxChunkSize)
        {
            var chunks = new List<string>();
            
            if (input.Length <= maxChunkSize)
            {
                chunks.Add(input);
                return chunks;
            }

            var sentences = input.Split(new[] { ". ", ".\n", "!\n", "?\n" }, StringSplitOptions.None);
            var currentChunk = "";
            
            foreach (var sentence in sentences)
            {
                var testChunk = currentChunk + (currentChunk.Length > 0 ? ". " : "") + sentence;
                
                if (testChunk.Length > maxChunkSize && currentChunk.Length > 0)
                {
                    chunks.Add(currentChunk);
                    currentChunk = sentence;
                }
                else
                {
                    currentChunk = testChunk;
                }
            }
            
            if (currentChunk.Length > 0)
            {
                chunks.Add(currentChunk);
            }
            
            return chunks;        }

        // Original processing logic with enhanced safety
        private static bool ProcessNormalContent(string input)
        {
            input = " " + input;
            Regex mathRegex = new Regex(@"<math [\S\s]*?>[\S\s]*?<\/math>");

            string[] arr1 = mathRegex.Split(input);
            string[] arr2 = mathRegex.Matches(input).Cast<Match>().Select(m => m.Value).ToArray();
            if (arr2.Length == 0) return false;

            List<string> returnArr = arr1.SelectMany((el, index) => new[] { el, arr2.ElementAtOrDefault(index) }).ToList();
            returnArr.RemoveAll(string.IsNullOrEmpty);
            
            // ── enhanced adaptive throttling ────────────────────────────────
            Word.Application app = Globals.Main.Application;
            int existingEq = CountExistingEquations(app.ActiveDocument);
            int incomingEq = arr2.Length;
            int contentLength = input.Length;

            // More conservative thresholds based on combined factors
            int CHUNK, DELAY;
            
            if (existingEq > 200 || incomingEq > 30 || contentLength > 30000)
            {
                CHUNK = 2;   // Very small chunks for heavy load
                DELAY = 600; // Long delays
            }
            else if (existingEq > 100 || incomingEq > 15 || contentLength > 15000)
            {
                CHUNK = 3;   // Small chunks
                DELAY = 400; // Medium delays
            }
            else if (existingEq > 50 || incomingEq > 8 || contentLength > 8000)
            {
                CHUNK = 5;   // Medium chunks
                DELAY = 250; // Short delays
            }
            else
            {
                CHUNK = 10;  // Default chunks
                DELAY = 150; // Minimal delays
            }

            System.Diagnostics.Debug.WriteLine($"Paste strategy: {existingEq} existing, {incomingEq} incoming, {contentLength} chars → chunk={CHUNK}, delay={DELAY}ms");

            bool oldUpd = app.ScreenUpdating;
            app.ScreenUpdating = false;

            try
            {
                int counter = 0;
                for (int i = returnArr.Count - 1; i >= 0; i--)
                {
                    Paste(returnArr[i]);
                    counter++;

                    // Always add a small pause after each paste for stability
                    System.Threading.Thread.Sleep(50);
                    System.Windows.Forms.Application.DoEvents();

                    if (counter >= CHUNK)
                    {
                        System.Threading.Thread.Sleep(DELAY);
                        System.Windows.Forms.Application.DoEvents();
                        counter = 0;
                    }
                }
                
                return true;
            }
            finally
            {
                app.ScreenUpdating = oldUpd;            }
        }

        public static bool PasteEquation()
        {
            try
            {
                Range currentRange = Globals.Main.Application.Selection.Range;
                currentRange.Text = " ";
                currentRange.SetRange(currentRange.End, currentRange.End);
                currentRange.Select();

                string clipboardText = Clipboard.GetText();

                bool returnVal = true;
                if (clipboardText == string.Empty || !SplitAndPaste(clipboardText)) 
                    returnVal = false;
                else 
                    Clipboard.SetText(clipboardText);

                currentRange.SetRange(currentRange.Start - 1, currentRange.Start - 1);
                currentRange.Delete(WdUnits.wdCharacter, 2);
                return returnVal;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"PasteEquation error: {ex.Message}");
                
                // Try to restore clipboard content if possible
                try
                {
                    string clipboardText = Clipboard.GetText();
                    if (!string.IsNullOrEmpty(clipboardText))
                        Clipboard.SetText(clipboardText);
                }
                catch { }
                
                return false;
            }
        }
    }

    public class KeyboardHook {
        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod,
            uint dwThreadId);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string lpModuleName);

        public delegate int LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);
        private static LowLevelKeyboardProc _proc = HookCallback;
        private static IntPtr _hookID = IntPtr.Zero;


        private const int WH_KEYBOARD = 2;
        private const int HC_ACTION = 0;
        public static void SetHook()
        {
#pragma warning disable 618
            _hookID = SetWindowsHookEx(WH_KEYBOARD, _proc, IntPtr.Zero, (uint)AppDomain.GetCurrentThreadId());
#pragma warning restore 618

        }

        public static void ReleaseHook()
        {
            UnhookWindowsHookEx(_hookID);
        }
        private static int HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode < 0)
            {
                return (int)CallNextHookEx(_hookID, nCode, wParam, lParam);
            }
            else
            {

                if (nCode == HC_ACTION)
                {
                    Keys keyData = (Keys)wParam;

                    if ((BindingFunctions.IsKeyDown(Keys.ControlKey) == true)
                    && (BindingFunctions.IsKeyDown(keyData) == true) && (keyData == Keys.V))
                    {
                        if (WordFunctions.PasteEquation()) return 1;
                    }

                }
                return (int)CallNextHookEx(_hookID, nCode, wParam, lParam);
            }
        }
    }

    public class BindingFunctions
    {
        [DllImport("user32.dll")]
        static extern short GetKeyState(int nVirtKey);

        public static bool IsKeyDown(Keys keys)
        {
            return (GetKeyState((int)keys) & 0x8000) == 0x8000;
        }

    }
}